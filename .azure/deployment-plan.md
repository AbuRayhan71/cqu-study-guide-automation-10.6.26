# CQU Study Guide Automator - Azure Deployment Plan

## Overview
- **Project**: CQU Study Guide Automation Web App
- **Mode**: MODIFY (existing FastAPI + React/HTML application)
- **Target Platform**: Azure App Service
- **Current Stack**: Python FastAPI, HTML/CSS/JS frontend, DOCX processing

## Phase 1: Planning

### Phase 1 - Step 1: Analyze Workspace ✓
- **Status**: COMPLETE
- **Mode**: MODIFY — existing FastAPI app with static assets
- **Key findings**:
  - FastAPI application (app/main.py)
  - Static HTML/CSS/JS frontend (app/static/)
  - Python dependencies (requirements.txt)
  - Uses Groq API for AI features
  - `.env` configuration with API keys

### Phase 1 - Step 2: List Azure Services
- **Status**: COMPLETE
- **Services required**:
  - Azure App Service (Python 3.12 runtime)
  - Storage Account (for uploaded/output DOCX files)
  - Application Insights (monitoring)
  - Key Vault (for API key management)

### Phase 1 - Step 3: Select IaC Format
- **Status**: COMPLETE
- **Choice**: Bicep

### Phase 1 - Step 4: Select Recipe
- **Status**: COMPLETE
- **Recipe**: Web App + Storage + Application Insights

### Phase 1 - Step 5: Infrastructure Configuration
- **Status**: PENDING
- **To be configured**:
  - [ ] Azure Subscription ID
  - [ ] Resource Group name
  - [ ] Region
  - [ ] App Service Plan tier
  - [ ] Storage account name
  - [ ] Managed Identity (yes/no)

### Phase 1 - Step 6: Finalize & Approve Plan
- **Status**: PENDING USER APPROVAL

## Phase 2: Code Generation (After Approval)
- [ ] Generate Bicep infrastructure code
- [ ] Create azure.yaml for deployment
- [ ] Generate Dockerfile (containerize FastAPI app)
- [ ] Update requirements.txt if needed

## Phase 3: Validation
- [ ] Run azure-validate to check configuration

## Phase 4: Deployment
- [ ] Run azure-deploy to provision resources and deploy app

---

## Ready for Configuration

When you're ready to add Azure resource group details, please provide:
1. Azure Subscription ID
2. Resource Group name
3. Region (e.g., australiaeast, eastus)
4. App Service Plan tier (B1, B2, B3)
5. Storage Account name (3-24 chars, lowercase, globally unique)
