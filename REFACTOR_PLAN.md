# Technical Refactor Plan: DailyWD-UK V2.0

This document outlines the steps to move the DailyWD-UK pipeline into a professional, server-ready state with a user-friendly management interface.

## 🟢 PHASE 1: Core Optimization (COMPLETED)
1. **Centralized Data & Settings Loader:** Consolidated logic into `data_loader.py`.
2. **Database-Backed Summaries:** Migrated Excel to SQLite `daily_summary` table.
3. **Automated Settings Validator:** Pre-flight check added to catch Excel errors.
4. **SQL-Level Filtering:** Optimized fetching for millions of rows.
5. **Fully Centralized Paths:** Single source of truth for all file paths.
6. **Decoupled Parallelism:** Parallel file generation + Sequential Outlook sending.
7. **Automated Disk Cleanup:** Storage maintenance.

## 🔵 PHASE 2: Full-Stack Modernization (COMPLETED)
**Goal:** Replace basic scripts with an enterprise-grade web application.

### 1. Backend Service (FastAPI)
* **Status:** ✅ DONE
* Implemented `api_server.py` with REST and WebSocket support.
* Developed "Smart Merge" UPSERT logic for persistent roster updates.
* Automated "Ownership Matrix" lookup and manager syncing.

### 2. Frontend Dashboard (Next.js 14)
* **Status:** ✅ DONE
* High-density UI for managing large teams.
* Visual Progress Stepper for non-technical observability.
* Granular multi-select filters (Role, Region, SubRegion).

### 3. Intelligent Import
* **Status:** ✅ DONE
* Column mapping for non-standard Excel headers (e.g., "Email ID (Official)").
* Metadata support for Week-offs and metadata preservation.

---
**Updated on:** 12 May 2026 (Enterprise Suite Finalized)
