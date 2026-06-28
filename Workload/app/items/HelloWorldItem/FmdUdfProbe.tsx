import React, { useCallback, useEffect, useState } from "react";
import { Button, Spinner, Text, Badge } from "@fluentui/react-components";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { callAcquireFrontendAccessToken } from "../../controller/AuthenticationController";

/**
 * FMD dashboard (native FET frontend) — proves AND demos the host-brokered auth path.
 *
 * On open it asks the Fabric HOST for a Power BI–audience token (acquireFrontendAccessToken — no
 * MSAL, no popup, no iframe handoff) and calls the Entra-protected FMD UDF directly: dashboard_kpis,
 * list_sources (Domain → subject_area → dataset), lane_summary and list_jobs. Spike config hardcoded
 * to FMD-DEV; productionize via workload env/config.
 */
const UDF_BASE =
  "https://32ff430a586143028cf3bc7d75cc3093.z32.userdatafunctions.fabric.microsoft.com" +
  "/v1/workspaces/32ff430a-5861-4302-8cf3-bc7d75cc3093" +
  "/userDataFunctions/42bbfd68-e376-4209-9a94-30d2922b2483";
const UDF_SCOPE = "https://analysis.windows.net/powerbi/api/UserDataFunction.Execute.All";

type Kpis = Record<string, number>;
interface SourceRow {
  domain: string | null; subject_area: string | null; asset: string | null;
  lane: string | null; readiness_code: string | null;
}
interface LaneRow {
  lane: string | null; asset_count: number; ready_count: number;
  pending_count: number; blocked_count: number;
}
interface JobRow {
  asset_name: string | null; acquisition_type: string | null;
  status: string | null; started_utc: string | null;
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

type Tone = "success" | "warning" | "danger" | "informative";
const readinessTone = (code: string | null): Tone => {
  const c = (code || "").toUpperCase();
  if (c === "READY") return "success";
  if (c === "PENDING") return "warning";
  if (c.includes("NO_CONNECTION") || c.includes("FAILED") || c.includes("BLOCK")) return "danger";
  return "informative";
};
const statusTone = (s: string | null): Tone => {
  const c = (s || "").toUpperCase();
  if (c === "SUCCEEDED") return "success";
  if (c === "FAILED" || c === "QUARANTINED") return "danger";
  if (c === "RUNNING" || c === "STARTED") return "warning";
  return "informative";
};

const KPI_CARDS: [string, string][] = [
  ["sources", "Sources"], ["assets", "Assets"], ["ready", "Ready"],
  ["pending", "Pending"], ["blocked", "Blocked"], ["quarantined", "Quarantined"],
];
const th: React.CSSProperties = { textAlign: "left", padding: "6px 10px", borderBottom: "2px solid #ddd" };
const td: React.CSSProperties = { padding: "6px 10px", borderBottom: "1px solid #eee", fontSize: 13 };

function LanesTile({ lanes }: { lanes: LaneRow[] }) {
  return (
    <div style={{ marginTop: 24 }}>
      <Text size={500} weight="semibold">Lanes</Text>
      <table style={{ borderCollapse: "collapse", marginTop: 8, minWidth: 460 }}>
        <thead>
          <tr><th style={th}>Lane</th><th style={th}>Assets</th><th style={th}>Ready</th>
            <th style={th}>Pending</th><th style={th}>Blocked</th></tr>
        </thead>
        <tbody>
          {lanes.map((l, i) => (
            <tr key={i}>
              <td style={td}><Badge size="small" appearance="outline">{l.lane}</Badge></td>
              <td style={td}>{l.asset_count}</td>
              <td style={td}>{l.ready_count}</td>
              <td style={td}>{l.pending_count}</td>
              <td style={td}>{l.blocked_count}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function JobsTile({ jobs }: { jobs: JobRow[] }) {
  return (
    <div style={{ marginTop: 24 }}>
      <Text size={500} weight="semibold">Recent jobs</Text>
      <table style={{ borderCollapse: "collapse", marginTop: 8, minWidth: 560 }}>
        <thead>
          <tr><th style={th}>Asset</th><th style={th}>Lane</th><th style={th}>Status</th>
            <th style={th}>Started</th></tr>
        </thead>
        <tbody>
          {jobs.slice(0, 8).map((j, i) => (
            <tr key={i}>
              <td style={td}>{j.asset_name}</td>
              <td style={td}>{j.acquisition_type}</td>
              <td style={td}><Badge size="small" color={statusTone(j.status)}>{j.status || "—"}</Badge></td>
              <td style={td}>{j.started_utc ? j.started_utc.slice(0, 19).replace("T", " ") : "—"}</td>
            </tr>
          ))}
          {jobs.length === 0 && <tr><td style={td} colSpan={4}>No runs yet.</td></tr>}
        </tbody>
      </table>
    </div>
  );
}

export function FmdUdfProbe({ workloadClient }: { workloadClient: WorkloadClientAPI }) {
  const [kpis, setKpis] = useState<Kpis | null>(null);
  const [sources, setSources] = useState<SourceRow[]>([]);
  const [lanes, setLanes] = useState<LaneRow[]>([]);
  const [jobs, setJobs] = useState<JobRow[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState<boolean>(true);

  const load = useCallback(async (): Promise<void> => {
    setLoading(true);
    setError(null);
    try {
      // The whole point: the host hands us the token — we never touch MSAL.
      const { token } = await callAcquireFrontendAccessToken(workloadClient, UDF_SCOPE);
      const [k, s, l, j] = await Promise.all([
        invokeUdf<Kpis>(token, "dashboard_kpis"),
        invokeUdf<SourceRow[]>(token, "list_sources"),
        invokeUdf<LaneRow[]>(token, "lane_summary"),
        invokeUdf<JobRow[]>(token, "list_jobs"),
      ]);
      setKpis(k);
      setSources(Array.isArray(s) ? s : []);
      setLanes(Array.isArray(l) ? l : []);
      setJobs(Array.isArray(j) ? j : []);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }, [workloadClient]);

  useEffect(() => { void load(); }, [load]);

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
        <Button size="small" onClick={() => { void load(); }} disabled={loading}>Refresh</Button>
        {loading && <Spinner size="tiny" label="Loading via host-brokered token…" />}
      </div>

      {error && <p role="alert" style={{ color: "#b10e1c", marginTop: 12 }}>⚠️ {error}</p>}

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
                      <Badge size="small" color={readinessTone(r.readiness_code)}>{r.readiness_code || "—"}</Badge>
                    </div>
                  ))}
                </div>
              ))}
            </div>
          ))}
        </div>
      )}

      {lanes.length > 0 && <LanesTile lanes={lanes} />}
      {kpis && <JobsTile jobs={jobs} />}
    </div>
  );
}
