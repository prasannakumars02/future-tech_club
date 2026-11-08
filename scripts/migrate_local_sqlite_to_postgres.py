import getpass
import sys
from pathlib import Path

from sqlalchemy import Integer, MetaData, create_engine, func, inspect, select, text, update
from sqlalchemy.exc import ArgumentError
from sqlalchemy.engine import make_url
from werkzeug.security import generate_password_hash

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from app import models
from app.extensions import db

SOURCE_DATABASE = ROOT / "instance" / "future_tech_innovision.db"


def postgres_url(value):
    url = make_url(value.strip())
    if url.drivername in {"postgres", "postgresql"}:
        url = url.set(drivername="postgresql+psycopg")
    if url.drivername != "postgresql+psycopg":
        raise ValueError("The destination must be a PostgreSQL connection URL.")
    return url


def reflected_source_tables(engine):
    metadata = MetaData()
    metadata.reflect(bind=engine)
    return metadata


def copy_database(source_engine, destination_engine, admin_username, admin_password):
    source_metadata = reflected_source_tables(source_engine)
    app_tables = db.metadata.sorted_tables
    app_table_names = {table.name for table in app_tables}
    source_table_names = set(source_metadata.tables)
    destination_table_names = set(inspect(destination_engine).get_table_names())

    missing_source = app_table_names - source_table_names
    missing_destination = app_table_names - destination_table_names
    if missing_source:
        raise RuntimeError(
            "The local database is missing application tables: "
            + ", ".join(sorted(missing_source))
        )
    if missing_destination:
        raise RuntimeError(
            "The hosted database is missing application tables. "
            "Deploy the application successfully before migrating."
        )

    user_table = models.User.__table__
    with source_engine.connect() as source_connection:
        source_counts = {
            table.name: source_connection.execute(
                select(func.count()).select_from(source_metadata.tables[table.name])
            ).scalar_one()
            for table in app_tables
        }
        source_admin = source_connection.execute(
            select(user_table.c.id).where(
                user_table.c.username == admin_username,
                user_table.c.role == "admin",
            )
        ).scalar_one_or_none()
    if source_admin is None:
        raise RuntimeError(
            "The local database has no admin account matching "
            "BOOTSTRAP_ADMIN_USERNAME in Render."
        )

    print(f"Local application tables: {len(app_tables)}")
    print(f"Local application records: {sum(source_counts.values())}")
    print("The hosted application tables will be replaced by these local records.")
    print("This operation cannot be undone. Ensure you have a separate backup.")
    if input('Type "REPLACE" to continue: ').strip() != "REPLACE":
        print("Migration cancelled; the hosted database was not changed.")
        return False

    with destination_engine.begin() as connection:
        preparer = destination_engine.dialect.identifier_preparer
        quoted_tables = ", ".join(
            preparer.quote(table.name) for table in app_tables
        )
        connection.execute(
            text(f"TRUNCATE TABLE {quoted_tables} RESTART IDENTITY CASCADE")
        )

        for target_table in app_tables:
            source_table = source_metadata.tables[target_table.name]
            source_columns = {column.name for column in source_table.columns}
            target_column_names = {column.name for column in target_table.columns}
            shared_columns = [
                column.name
                for column in target_table.columns
                if column.name in (source_columns & target_column_names)
            ]
            missing_required = [
                column.name
                for column in target_table.columns
                if column.name not in source_columns
                and not column.nullable
                and column.default is None
                and column.server_default is None
                and not (
                    column.primary_key
                    and isinstance(column.type, Integer)
                    and column.autoincrement is not False
                )
            ]
            if missing_required:
                raise RuntimeError(
                    f"Local table {target_table.name} lacks required columns: "
                    + ", ".join(missing_required)
                )

            with source_engine.connect() as source_connection:
                rows = source_connection.execute(select(source_table)).mappings()
                batch = [{name: row[name] for name in shared_columns} for row in rows]
            if batch:
                connection.execute(target_table.insert(), batch)

        for target_table in app_tables:
            for column in target_table.primary_key.columns:
                if not isinstance(column.type, Integer) or column.autoincrement is False:
                    continue
                sequence_name = connection.execute(
                    text("SELECT pg_get_serial_sequence(:table_name, :column_name)"),
                    {"table_name": target_table.name, "column_name": column.name},
                ).scalar_one_or_none()
                if not sequence_name:
                    continue
                maximum = connection.execute(
                    select(func.max(column)).select_from(target_table)
                ).scalar_one()
                connection.execute(
                    text(
                        "SELECT setval(CAST(:sequence_name AS regclass), "
                        ":last_value, :is_called)"
                    ),
                    {
                        "sequence_name": sequence_name,
                        "last_value": maximum if maximum is not None else 1,
                        "is_called": maximum is not None,
                    },
                )

        destination_admin = connection.execute(
            select(user_table.c.id).where(
                user_table.c.username == admin_username,
                user_table.c.role == "admin",
            )
        ).scalar_one_or_none()
        if destination_admin is None:
            raise RuntimeError("The migrated admin account could not be found.")
        connection.execute(
            update(user_table)
            .where(user_table.c.id == destination_admin)
            .values(password_hash=generate_password_hash(admin_password))
        )

    print("Migration completed successfully.")
    print(f"Imported application records: {sum(source_counts.values())}")
    return True


def main():
    if not SOURCE_DATABASE.is_file():
        raise SystemExit(f"Local database not found: {SOURCE_DATABASE}")

    admin_username = input(
        "Enter the BOOTSTRAP_ADMIN_USERNAME configured in Render: "
    ).strip()
    if not admin_username:
        raise SystemExit("Admin username must not be empty.")
    admin_password = getpass.getpass(
        "Enter the BOOTSTRAP_ADMIN_PASSWORD configured in Render (hidden): "
    )
    confirmed_password = getpass.getpass("Confirm the admin password (hidden): ")
    if admin_password != confirmed_password:
        raise SystemExit("Passwords do not match.")
    if len(admin_password) < 12:
        raise SystemExit("Production admin password must be at least 12 characters.")

    destination_value = getpass.getpass(
        "Paste the Neon PostgreSQL connection string (input is hidden): "
    )
    try:
        destination_url = postgres_url(destination_value)
    except (ArgumentError, ValueError) as error:
        raise SystemExit(
            "Invalid PostgreSQL connection URL. Copy the full connection string from Neon."
        ) from error
    finally:
        destination_value = None

    source_engine = create_engine(
        f"sqlite:///{SOURCE_DATABASE.as_posix()}",
        connect_args={"check_same_thread": False},
    )
    destination_engine = create_engine(destination_url, hide_parameters=True)

    try:
        with destination_engine.connect() as connection:
            if connection.exec_driver_sql("SELECT 1").scalar_one() != 1:
                raise RuntimeError("Could not verify the PostgreSQL connection.")
        copy_database(
            source_engine, destination_engine, admin_username, admin_password
        )
    finally:
        source_engine.dispose()
        destination_engine.dispose()
        admin_password = None
        confirmed_password = None


if __name__ == "__main__":
    main()
