.PHONY: setup check data spec up demo demo-bundled demo-inline kill-demo psql logs restart-workers down reset
setup:  ## create .env from the template (never overwrites an existing one)
	@test -f .env && echo ".env already exists" || (cp .env.example .env && echo "Created .env: put your Gemini key in it")
check:  ## prove Docker, the sample data and (if set) the key work
	./run.sh check.py
data:   ## rebuild data/ from data/datagen/generate.py (same fonts and library versions every time)
	docker run --rm -v "$$PWD/data:/data" -w /data/datagen python:3.12-slim sh -c \
	  "apt-get update -qq >/dev/null && apt-get install -y -qq fonts-dejavu-core >/dev/null && \
	   pip install -q --root-user-action=ignore reportlab==4.4.4 openpyxl==3.1.5 pillow==12.3.0 && python generate.py"
spec:   ## rebuild docs/ClearRail_Use_Cases.pdf from its HTML source (uses the installed Google Chrome)
	uv run -q --with playwright==1.55.0 python docs/build_spec.py

# --- steps 4 and 5: a database, an API, a screen and a queue -----------------------------------------
up:     ## build and start Postgres, Redis, the API + screen and both worker pools
	@[ -f .env ] || cp .env.example .env
	docker compose up --build -d
	@echo "Screen + API  http://localhost:8020   (Swagger at /docs)   Postgres on localhost:5434"
demo:   ## upload the 32 UC1 payments as one batch through the API, and watch the workers check them
	docker compose run --rm api python ladder/step5/load.py
demo-bundled:  ## the same for the 15 bundled payments (reads new scans with the model once)
	docker compose run --rm api python ladder/step5/load.py --bundled
demo-inline:  ## step 4's way, for comparison: read and check each payment inline, no workers
	docker compose run --rm api python ladder/step4/load.py
kill-demo:  ## on an empty database: slow reads to 1 s, upload all 47 payments, SIGKILL the reader mid-batch
	DEMO_READ_DELAY=1 docker compose up -d worker-reader
	@( sleep 12 ; echo "$$(date +%T)  killing worker-reader (SIGKILL), mid-batch" ; \
	   docker compose kill -s SIGKILL worker-reader >/dev/null 2>&1 ; sleep 3 ; \
	   echo "$$(date +%T)  starting worker-reader again" ; \
	   DEMO_READ_DELAY=1 docker compose up -d worker-reader >/dev/null 2>&1 ) &
	docker compose run --rm api python ladder/step5/load.py --all
	docker compose up -d worker-reader
psql:   ## open a SQL prompt on the database
	docker compose exec postgres psql -U clearrail
logs:   ## follow both worker pools
	docker compose logs -f --tail=30 worker-reader worker-checker
restart-workers:  ## pick up a code change in the workers
	docker compose restart worker-reader worker-checker
down:   ## stop the stack (the database is kept)
	docker compose down
reset:  ## stop the stack and DELETE the database (uploaded files stay in data/storage/)
	docker compose down -v
