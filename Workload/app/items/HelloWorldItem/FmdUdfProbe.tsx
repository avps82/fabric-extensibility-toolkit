import React, { useCallback, useEffect, useState } from "react";
import { Button, Spinner, Text, Badge } from "@fluentui/react-components";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { callAcquireFrontendAccessToken } from "../../controller/AuthenticationController";

/**
 * FMD dashboard (native FET frontend) — proves AND demos the host-brokered auth path.
 *
 * On open it asks the Fabric HOST for a Power BI–audience token (acquireFrontendAccessToken — no
 * MSAL, no popup, no iframe handoff) and calls the Entra-protected FMD UDF directly: dashboard_kpis
 * for the headline counts and list_sources for the Domain → subject_area → dataset tree. Spike
 * config hardcoded to FMD-DEV; productionize via workload env/config.
 */
const UDF_BASE =
  "https://32ff430a586143028cf3bc7d75cc3093.z32.userdatafunctions.fabric.microsoft.com" +
  "/v1/workspaces/32ff430a-5861-4302-8cf3-bc7d75cc3093" +
  "/userDataFunctions/42bbfd68-e376-4209-9a94-30d2922b2483";
const UDF_SCOPE = "https://analysis.windows.net/powerbi/api/UserDataFunction.Execute.All";

type Kpis = Record<string, number>;
interface SourceRow {
  domain: string | null;
  subject_area: string | null;
  asset: string | null;
  lane: string | null;
  readiness_code: string | null;
}

async function invokeUdf<T>(token: string, fn: string): Promise<T> {
  const r = await fetch(`${UDF_BASE}/functions/${fn}/invoke`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: "{}",
  });
  const body = await r.json();
  if (body.status !== "Succeeded") {
    throw new Error(`${fn}: ${body.status} ${JSON.stringify(body.errors)}`);
  }
  return body.output as T;
}

const tone = (code: string | null): "success" | "warning" | "danger" | "informative" => {
  const c = (code || "").toUpperCase();
  if (c === "READY") return "success";
  if (c === "PENDING") return "warning";
  if (c.includes("NO_CONNECTION") || c.includes("FAILED") || c.includes("BLOCK")) return "danger";
  return "informative";
};

const KPI_CARDS: [string, string][] = [
  ["sources", "Sources"], ["assets", "Assets"], ["ready", "Ready"],
  ["pending", "Pending"], ["blocked", "Blocked"], ["quarantined", "Quarantined"],
];

export function FmdUdfProbe({ workloadClient }: { workloadClient: WorkloadClientAPI }) {
  const [kpis, setKpis] = useState<Kpis | null>(null);
  const [sources, setSources] = useState<SourceRow[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState<boolean>(true);

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      // The whole point: the host hands us the token — we never touch MSAL.
      const { token } = await callAcquireFrontendAccessToken(workloadClient, UDF_SCOPE);
      const [k, s] = await Promise.all([
        invokeUdf<Kpis>(token, "dashboard_kpis"),
        invokeUdf<SourceRow[]>(token, "list_sources"),
      ]);
      setKpis(k);
      setSources(Array.isArray(s) ? s : []);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }, [workloadClient]);

  useEffect(() => {
    void load();
  }, [load]);

  // Group sources: Domain -> subject_area -> datasets
  const domains = new Map<string, Map<string, SourceRow[]>>();
  for (const r of sources) {
    const d = r.domain || "(unknown)";
    const sa = r.subject_area || "(none)";
    const m = domains.get(d) || new Map<string, SourceRow[]>();
    (m.get(sa) || m.set(sa, []).get(sa)!).push(r);
    domains.set(d, m);
  }

  return (
    <div style={{ padding: 20, maxWidth: 1100 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
        <Text size={600} weight="bold">FMD — Bronze Ingestion Control Plane</Text>
        <Button size="small" onClick={() => void load()} disabled={loading}>Refresh</Button>
        {loading && <Spinner size="tiny" label="Loading via host-brokered token…" />}
      </div>

      {error && (
        <p role="alert" style={{ color: "#b10e1c", marginTop: 12 }}>⚠️ {error}</p>
      )}

      {kpis && (
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginTop: 16 }}>
          {KPI_CARDS.map(([key, label]) => (
            <div key={key} style={{
              minWidth: 130, padding: 16, borderRadius: 8, background: "#f5f5f5",
              border: "1px solid #e0e0e0",
            }}>
              <Text size={800} weight="bold">{kpis[key] ?? 0}</Text>
              <div><Text size={200}>{label}</Text></div>
            </div>
          ))}
        </div>
      )}

      {domains.size > 0 && (
        <div style={{ marginTop: 24 }}>
          <Text size={500} weight="semibold">Sources — Domain → Subject area → Dataset</Text>
          {[...domains.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([domain, areas]) => (
            <div key={domain} style={{ marginTop: 12 }}>
              <Text weight="bold">{domain.replace("__DEMO__", "")}</Text>
              {[...areas.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([sa, rows]) => (
                <div key={sa} style={{ marginLeft: 16, marginTop: 4 }}>
                  <Text size={300} style={{ color: "#616161", textTransform: "uppercase" }}>{sa}</Text>
                  {rows.map((r, i) => (
                    <div key={i} style={{ marginLeft: 16, display: "flex", gap: 8, alignItems: "center", padding: "2px 0" }}>
                      <Text size={300}>{r.asset}</Text>
                      <Badge size="small" appearance="outline">{r.lane}</Badge>
                      <Badge size="small" color={tone(r.readiness_code)}>{r.readiness_code || "—"}</Badge>
                    </div>
                  ))}
                </div>
              ))}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
