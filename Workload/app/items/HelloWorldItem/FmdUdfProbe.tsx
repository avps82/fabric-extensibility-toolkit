import React, { useState } from "react";
import { Button, Spinner, Text } from "@fluentui/react-components";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { callAcquireFrontendAccessToken } from "../../controller/AuthenticationController";

/**
 * FMD SPIKE — proves the durable native-frontend auth path.
 *
 * The Fabric host brokers a Power BI–audience token via `acquireFrontendAccessToken`
 * (callAcquireFrontendAccessToken), and we call the Entra-protected FMD onboarding UDF directly with
 * it. No `@azure/msal-browser`, no popup, no iframe handoff — i.e. none of the failure modes the
 * rayfin app hit (monitor_window_timeout / interaction_in_progress / popup handoff). If this button
 * renders the KPIs inside the Fabric portal, the FET frontend is the right home for the FMD UI.
 *
 * Spike config is hardcoded to FMD-DEV for the proof; productionize via the workload env/config.
 */
const UDF_BASE =
  "https://32ff430a586143028cf3bc7d75cc3093.z32.userdatafunctions.fabric.microsoft.com" +
  "/v1/workspaces/32ff430a-5861-4302-8cf3-bc7d75cc3093" +
  "/userDataFunctions/42bbfd68-e376-4209-9a94-30d2922b2483";
const UDF_SCOPE = "https://analysis.windows.net/powerbi/api/UserDataFunction.Execute.All";

export function FmdUdfProbe({ workloadClient }: { workloadClient: WorkloadClientAPI }) {
  const [status, setStatus] = useState<string>("");
  const [kpis, setKpis] = useState<Record<string, unknown> | null>(null);
  const [busy, setBusy] = useState<boolean>(false);

  async function probe(): Promise<void> {
    setBusy(true);
    setKpis(null);
    setStatus("acquiring host-brokered token…");
    try {
      // THE point of the spike: the host returns the token — we never instantiate MSAL.
      const { token } = await callAcquireFrontendAccessToken(workloadClient, UDF_SCOPE);
      setStatus("calling FMD UDF dashboard_kpis…");
      const r = await fetch(`${UDF_BASE}/functions/dashboard_kpis/invoke`, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: "{}",
      });
      const body = await r.json();
      if (body.status !== "Succeeded") {
        throw new Error(`UDF status ${body.status}: ${JSON.stringify(body.errors)}`);
      }
      setKpis(body.output);
      setStatus("✅ host-brokered token → UDF OK");
    } catch (e) {
      setStatus(`❌ ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setBusy(false);
    }
  }

  return (
    <div style={{ padding: 16, borderTop: "1px solid #eee" }}>
      <Text weight="semibold">FMD spike — host-brokered token → UDF</Text>
      <div style={{ marginTop: 8 }}>
        <Button appearance="primary" onClick={probe} disabled={busy}>
          {busy ? <Spinner size="tiny" label="Working…" /> : "Call FMD UDF (dashboard_kpis)"}
        </Button>
      </div>
      {status && <p style={{ marginTop: 8 }}>{status}</p>}
      {kpis && (
        <pre style={{ marginTop: 8, background: "#f3f3f3", padding: 8 }}>
          {JSON.stringify(kpis, null, 2)}
        </pre>
      )}
    </div>
  );
}
