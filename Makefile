# SEPAD — operaciones fundamentales
# Uso: make <objetivo>   (make help para ver la lista)

PORT ?= 8080
IMAGE ?= sepad
URL := http://localhost:$(PORT)/

.DEFAULT_GOAL := help
.PHONY: help install setup start debug reset check crawl docker-build docker-run clean

help: ## Muestra esta ayuda
	@echo "SEPAD - objetivos disponibles:"
	@grep -E '^[a-zA-Z_-]+:.*?## ' $(MAKEFILE_LIST) | awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-14s\033[0m %s\n", $$1, $$2}'
	@echo ""
	@echo "Variables: PORT=$(PORT)  IMAGE=$(IMAGE)   (ej.: make start PORT=3000)"

node_modules: package.json package-lock.json
	npm install
	@touch node_modules

install: node_modules ## Instala las dependencias de Node

data/sepad.seed.sqlite: node_modules database/Sepad.mdb scripts/migrate-mdb.mjs scripts/schema-extra.sql scripts/seed-demo.mjs
	npm run setup

setup: data/sepad.seed.sqlite ## Crea la base de datos SQLite con los datos de demostración

start: setup ## Levanta SEPAD en http://localhost:PORT/
	@echo "SEPAD en $(URL)  (Ctrl+C para detener)"
	PORT=$(PORT) node server/index.js

debug: setup ## Levanta SEPAD mostrando las trazas completas de los errores ASP
	PORT=$(PORT) SEPAD_DEBUG=1 node server/index.js

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
