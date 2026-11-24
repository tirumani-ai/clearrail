"""The database connection. DATABASE_URL is set by docker-compose.yml; the default reaches step 4's Postgres
from your own machine (port 5434), e.g. for a database tool."""
import os
from collections.abc import Iterator
from contextlib import contextmanager

from sqlalchemy import create_engine
from sqlalchemy.orm import Session, sessionmaker

DATABASE_URL = os.environ.get("DATABASE_URL", "postgresql+psycopg://clearrail:clearrail@localhost:5434/clearrail")

engine = create_engine(DATABASE_URL, pool_pre_ping=True)
SessionLocal = sessionmaker(bind=engine, expire_on_commit=False)


@contextmanager
def session_scope() -> Iterator[Session]:
    """Commit on success, roll back on any error: a payment is written whole or not at all."""
    session = SessionLocal()
    try:
        yield session
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()


def init_db() -> None:
    """Create the tables if they are missing. Safe to run any number of times."""
    from app.db.models import Base

    Base.metadata.create_all(engine)
