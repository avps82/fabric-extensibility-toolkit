# FMD spike — native frontend via the Fabric Extensibility Toolkit (host-brokered token)

**Goal:** prove the durable fix for the FMD onboarding UI's auth pain. The rayfin "Fabric Apps
(Preview)" app made the browser run its own MSAL → popup/iframe handoff failures
(`monitor_window_timeout`, `interaction_in_progress`, popup never returns). An FET **workload** asks
the **Fabric host** for the token (`workloadClient.auth.acquireFrontendAccessToken`) — no MSAL in our
code, no popup we manage. This spike calls the live FMD UDF with that host-brokered token.

## What was changed (this fork)
- `Workload/app/items/HelloWorldItem/FmdUdfProbe.tsx` — a button that calls
  `callAcquireFrontendAccessToken(workloadClient, "…/UserDataFunction.Execute.All")` then POSTs the
  FMD UDF `dashboard_kpis/invoke` and renders the KPIs. Spike config (UDF URL + scope) is hardcoded
  to FMD-DEV at the top of the file — change it for another tenant.
- `Workload/app/items/HelloWorldItem/HelloWorldItemDefaultView.tsx` — renders `<FmdUdfProbe>` in the
  item's default view.

## Run it (your machine + tenant — needs the browser/portal)
Prereqs (see the repo README): Node LTS, PowerShell 7, .NET SDK, Azure CLI.
1. **Setup + Entra app:** `pwsh scripts/Setup/Setup.ps1` (creates the workload Entra app + env files).
   Then in Entra, on that app's **API permissions**, add the Power BI delegated scope
   **`UserDataFunction.Execute.All`** (so `acquireFrontendAccessToken` can request it for the UDF).
2. **Resume the FMD-DEV capacity** (the UDF needs it live): `az fabric capacity resume …` (Gmail-account
   subscription — see the FMD repo `docs/DEPLOY_STATUS.md`).
3. **Enable Fabric Developer Mode** in the tenant (Admin portal → Tenant settings) + your dev workspace.
4. **Run:** `npm install` in `Workload/`, then `npm run start:devServer` and (separate shell)
   `npm run start:devGateway` (or `scripts/Run/StartDevGateway.ps1`).
5. In the Fabric portal (Dev Mode), create a **HelloWorld** item from this workload, open it, and click
   **“Call FMD UDF (dashboard_kpis)”**.

## Pass / fail
- **PASS:** the button renders the KPI JSON (`{sources, assets, ready, …}`) — the host brokered a
  Power BI-audience token and the UDF accepted it, with **no popup and no `interaction_in_progress`**.
  → FET is the right home for the FMD UI; port the app's pages (Dashboard/Sources/Lanes/Jobs +
  onboarding form) into workload items, using `acquireFrontendAccessToken` for every UDF call.
- **FAIL (consent):** if the first call throws a consent error, call with `promptFullConsent: true`
  once (or have an admin grant tenant consent), then retry.

## Notes
- Build/typecheck were **not** run headless (installing this external repo's full dependency tree is
  gated); `npm install` + the dev server on your machine performs the real build.
- The UDF backend + scope are already proven (`deploy/probe_udf.sh` in the FMD repo → PROBE: PASS;
  delegated scope `UserDataFunction.Execute.All`). This spike only adds the **host-brokered token**
  acquisition that the embedded rayfin app couldn't do.
