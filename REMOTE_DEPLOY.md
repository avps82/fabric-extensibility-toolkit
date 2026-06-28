# FMD workload — publish for NATIVE hosting (no local devServer, any machine)

Goal: the `Org.FmdOnboarding` workload runs in Fabric on **any** machine with **no** `start:devServer`
/ DevGateway — because the frontend is hosted at a real URL and the workload is **published to the
tenant**. The app code is already proven (the dashboard works in dev mode); this only changes hosting.

Do it from the **home PC that already has the repo + node_modules working**. Once published, the demo
laptop just opens Fabric — nothing to install there.

## 1. Host the built frontend (GitHub Pages — free, no Azure)
Pages serves at a **subpath** (`/<repo>/`), so set `PUBLIC_PATH` (webpack now honors it) and a matching
`FRONTEND_URL`.

- Host URL will be: `https://avps82.github.io/fabric-extensibility-toolkit/`
- Edit `Workload/.env.prod` (create from `.env.template` if absent):
  ```
  WORKLOAD_HOSTING_TYPE=FERemote
  WORKLOAD_VERSION=1.0.0
  WORKLOAD_NAME=Org.FmdOnboarding
  ITEM_NAMES=HelloWorld
  FRONTEND_APPID=57c0c50a-c912-4b31-bca0-ad717b613937
  FRONTEND_URL=https://avps82.github.io/fabric-extensibility-toolkit/
  ENABLE_PLAYGROUND=false
  BACKEND_APPID=
  LOG_LEVEL=warn
  ENVIRONMENT_DISPLAY_NAME_SUFFIX=
  ```
- Build the FE:
  ```powershell
  cd Workload
  $env:PUBLIC_PATH = "/fabric-extensibility-toolkit/"
  npm run build:prod          # -> ../build/Frontend
  ```
- Publish `build/Frontend` to GitHub Pages (one-time): push its contents to a `gh-pages` branch, then
  repo **Settings → Pages → Source = gh-pages /(root)**. Confirm the bundle loads at the host URL.

## 2. Build + publish the workload package
The starter-kit's own tool reconfigures the manifest for remote hosting and rebuilds the `.nupkg`:
```powershell
pwsh scripts/Setup/SwitchToRemoteHosting.ps1   # answer FE-remote; FRONTEND_URL = the Pages URL above
```
This produces `build/Manifest/Org.FmdOnboarding.<ver>.nupkg` whose manifest points at the hosted FE
(no localhost). Then publish it (MS Learn — *Publish and manage your workload*):
- Fabric **Admin portal → Workloads → Publish** → upload the `.nupkg`.
- **Manage my tenant** → select the workload → **Consent** (org-wide) so users don't each consent.
- Tenant setting **"Users can see and work with additional workloads not validated by Microsoft"** must
  be on (already enabled).

## 3. Test (the payoff)
On a **clean machine** (no devServer): open Fabric → **Workload Hub** → `Org.FmdOnboarding` is there →
create a HelloWorld item → the FMD dashboard loads from the hosted FE via the host-brokered token.

## Notes / to validate together
- The Entra app `57c0c50a…` redirect URIs were created for `localhost:60006/close`; for remote it must
  also allow the Pages origin — add `https://avps82.github.io/fabric-extensibility-toolkit/close` (and
  the bare origin) as SPA redirect URIs on the app.
- UDF CORS already allows the Fabric portal origin (the call comes from the portal-hosted iframe).
- `SwitchToRemoteHosting.ps1` is 846 lines and asks several prompts (it also supports a backend — we're
  FE-only, so skip/blank backend fields). Run it interactively and we'll confirm the answers.
