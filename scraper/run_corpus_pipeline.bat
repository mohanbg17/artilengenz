@echo off
REM ==========================================================
REM   Corpus pipeline orchestrator
REM   Runs: fetch_so.py  ->  load_corpus.py  ->  embed_corpus.py
REM ==========================================================
REM Prerequisites:
REM   * cd to corpus-scraper directory
REM   * venv activated, requests + voyageai + pinecone-client installed
REM   * VOYAGE_API_KEY and PINECONE_API_KEY env vars set
REM   * Migration 02 already applied
REM ==========================================================

setlocal

set TARGET=%1
if "%TARGET%"=="" set TARGET=500

set OUTPUT=scraped_so.json

echo.
echo === Step 1: Scrape Stack Overflow (target: %TARGET%) ===
python fetch_so.py --target %TARGET% --output %OUTPUT%
if errorlevel 1 (
    echo [ERROR] fetch_so.py failed
    exit /b 1
)

echo.
echo === Step 2: Load into intel.corpus ===
python load_corpus.py --input %OUTPUT%
if errorlevel 1 (
    echo [ERROR] load_corpus.py failed
    exit /b 1
)

echo.
echo === Step 3: Embed into Pinecone (corpus_v1 namespace) ===
python embed_corpus.py
if errorlevel 1 (
    echo [ERROR] embed_corpus.py failed
    exit /b 1
)

echo.
echo === Pipeline complete ===
endlocal
