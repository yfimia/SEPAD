# SEPAD — operaciones fundamentales
# Uso: make <objetivo>   (make help para ver la lista)

PORT ?= 8080
IMAGE ?= sepad
URL := http://localhost:$(PORT)/
# Ruta absoluta del servidor: identifica solo los procesos de este SEPAD (no otros node)
SERVER := $(CURDIR)/server/index.js
PIDFILE := data/sepad.pid
LOGFILE := data/sepad.log

.DEFAULT_GOAL := help
.PHONY: help install setup start up stop restart status logs debug reset check crawl docker-build docker-run clean

help: ## Muestra esta ayuda
	@echo "SEPAD - objetivos disponibles:"
	@grep -E '^[a-zA-Z_-]+:.*?## ' $(MAKEFILE_LIST) | awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-14s\033[0m %s\n", $$1, $$2}'
	@echo ""
	@echo "Variables: PORT=$(PORT)  IMAGE=$(IMAGE)   (ej.: make up PORT=8082)"

node_modules: package.json package-lock.json
	npm install
	@touch node_modules

install: node_modules ## Instala las dependencias de Node

data/sepad.seed.sqlite: node_modules database/Sepad.mdb scripts/migrate-mdb.mjs scripts/schema-extra.sql scripts/seed-demo.mjs
	npm run setup

setup: data/sepad.seed.sqlite ## Crea la base de datos SQLite con los datos de demostración

start: setup ## Levanta SEPAD en primer plano (Ctrl+C para detener)
	@echo "SEPAD en $(URL)  (Ctrl+C para detener)"
	PORT=$(PORT) node $(SERVER)

up: setup ## Levanta SEPAD en segundo plano (log en data/sepad.log)
	@if pgrep -f "$(SERVER)" >/dev/null; then echo "SEPAD ya está corriendo (make status)"; exit 0; fi; \
	if lsof -nP -iTCP:$(PORT) -sTCP:LISTEN >/dev/null 2>&1; then echo "El puerto $(PORT) está ocupado por otro proceso. Elige otro, por ejemplo: make up PORT=8090"; exit 1; fi; \
	PORT=$(PORT) nohup node $(SERVER) > $(LOGFILE) 2>&1 & echo $$! > $(PIDFILE); \
	for i in $$(seq 1 50); do curl -s -o /dev/null $(URL) && break; sleep 0.2; done; \
	if curl -s -o /dev/null $(URL); then echo "SEPAD en $(URL)  (make stop para detener)"; \
	else echo "SEPAD no arrancó; revisa $(LOGFILE):"; tail -20 $(LOGFILE); exit 1; fi

stop: ## Detiene SEPAD (solo los procesos de este proyecto)
	@pids=$$(pgrep -f "$(SERVER)"); \
	if [ -z "$$pids" ]; then echo "SEPAD no está corriendo"; else kill $$pids && echo "SEPAD detenido"; fi; \
	: > $(PIDFILE)

restart: ## Reinicia SEPAD en segundo plano
	@$(MAKE) --no-print-directory stop
	@sleep 0.5
	@$(MAKE) --no-print-directory up PORT=$(PORT)

status: ## Indica si SEPAD está corriendo y en qué puerto
	@pids=$$(pgrep -f "$(SERVER)"); \
	if [ -z "$$pids" ]; then echo "SEPAD no está corriendo"; else \
	  for p in $$pids; do port=$$(lsof -nP -a -p $$p -iTCP -sTCP:LISTEN 2>/dev/null | awk 'NR>1 {sub(/.*:/, "", $$9); print $$9; exit}'); \
	  echo "SEPAD corriendo (pid $$p) en http://localhost:$$port/"; done; fi

logs: ## Muestra el log del servidor en segundo plano (Ctrl+C para salir)
	@tail -f $(LOGFILE)

debug: setup ## Levanta SEPAD en primer plano mostrando las trazas de los errores ASP
	PORT=$(PORT) SEPAD_DEBUG=1 node $(SERVER)

reset: node_modules ## Rehace la base de datos y los cursos de demostración desde cero
	npm run setup

check: node_modules ## Compila todas las páginas .asp y reporta errores
	npm run -s check

crawl: ## Recorre el sitio en ejecución como alumno, profesor y administrador
	node scripts/crawl.mjs $(URL) 600 alumno alumno
	node scripts/crawl.mjs $(URL) 600 profesor profesor
	node scripts/crawl.mjs $(URL) 600 administrador admin

docker-build: ## Construye la imagen Docker (requiere Docker Desktop abierto)
	docker build -t $(IMAGE) .

docker-run: docker-build ## Ejecuta la imagen Docker en http://localhost:PORT/
	docker run --rm -p $(PORT):8080 $(IMAGE)

clean: ## Borra la base de datos generada y la copia de los cursos
	rm -rf data Courses
