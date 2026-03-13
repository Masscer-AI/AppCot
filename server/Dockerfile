FROM python:3.12-slim

RUN pip install uv --no-cache-dir

WORKDIR /app

COPY pyproject.toml uv.lock* ./
RUN uv sync --frozen --no-dev 2>/dev/null || uv sync --no-dev

COPY . .

EXPOSE 8009

CMD ["uv", "run", "uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8009"]
