# The ClearRail lab image: Python and the libraries (step 4 adds the database, API and screen ones). The code is mounted at run time, so you can edit a file and run it again.
FROM python:3.12-slim
COPY --from=ghcr.io/astral-sh/uv:0.9.26 /uv /uvx /bin/
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 \
    UV_COMPILE_BYTECODE=1 UV_LINK_MODE=copy UV_NO_CACHE=1 \
    UV_PROJECT_ENVIRONMENT=/opt/venv PATH="/opt/venv/bin:$PATH"
WORKDIR /app
# The venv lives in /opt/venv, outside /app, so mounting the code over /app does not hide it.
COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev
