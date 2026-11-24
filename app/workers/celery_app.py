"""The queue: Celery on Redis. Two pools of workers take work from it (docker-compose.yml):

  reader   read.urgent, read.normal, read.bulk     one task per uploaded file (the model, where needed)
  checker  check.urgent, check.normal, check.bulk  one task per payment, once all its files are read (no model)

Results live in Postgres, never in the broker. The settings that make a crash lose nothing:
  task_acks_late              a task leaves the queue only when it FINISHES, not when a worker takes it
  task_reject_on_worker_lost  ...so the task of a worker that dies is put back
  visibility_timeout          how long Redis waits before handing an unfinished task to another worker.
                              Too short and a slow (but alive) read is handed out twice, which is safe here
                              (app/services/payments.py) but pays the model twice; too long and a crashed
                              worker's work waits that long. VISIBILITY_TIMEOUT, in seconds.
  queue_order_strategy        `priority`: a worker always empties .urgent before .normal before .bulk
                              (Redis's default is round robin, which would let a month-end pile delay an
                              urgent payment).
"""
import os

from celery import Celery

from app.db.models import PRIORITIES

celery_app = Celery("clearrail", broker=os.environ.get("REDIS_URL", "redis://localhost:6379/0"),
                    include=["app.workers.tasks"])
celery_app.conf.update(
    task_ignore_result=True,
    task_acks_late=True,
    task_reject_on_worker_lost=True,
    worker_prefetch_multiplier=1,  # a worker holds only the task it is working on, so urgent work is not stuck behind it
    broker_transport_options={"visibility_timeout": int(os.environ.get("VISIBILITY_TIMEOUT", "300")),
                              "queue_order_strategy": "priority"},
    broker_connection_retry_on_startup=True,
    task_default_queue="read.normal",
    task_serializer="json",
    accept_content=["json"],
    timezone="UTC",
)


def queue_for(pool: str, priority: str) -> str:
    return f"{pool}.{priority if priority in PRIORITIES else 'normal'}"


def queues(pool: str) -> str:
    """The -Q argument for a pool's workers, most urgent first."""
    return ",".join(queue_for(pool, p) for p in PRIORITIES)
