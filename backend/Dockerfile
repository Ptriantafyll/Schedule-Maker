FROM python:3.11.9-slim

# Pin exact uv version matching local development environment
COPY --from=ghcr.io/astral-sh/uv:0.10.12 /uv /bin/uv

# Set working directory
WORKDIR /app

# Configure Python and uv environment variables
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    UV_COMPILE_BYTECODE=1 \
    UV_LINK_MODE=copy \
    PATH="/app/.venv/bin:$PATH" \
    DATABASE_URL="sqlite:////app/data/hospital_schedule.db"

# Step 1: Copy dependency lockfiles first for optimal Docker layer caching
COPY pyproject.toml uv.lock ./

# Step 2: Install dependencies without installing the project source yet
RUN uv sync --frozen --no-dev --no-install-project

# Step 3: Copy application source code and seed data
COPY src/ ./src
COPY data/ ./data

# Step 4: Finalize sync to include our project root package
RUN uv sync --frozen --no-dev

# Ensure data directory exists for the mounted SQLite volume
RUN mkdir -p /app/data

# Expose internal API port
EXPOSE 8000

# Start Uvicorn bound to all interfaces
CMD ["uvicorn", "src.main:app", "--host", "0.0.0.0", "--port", "8000"]
