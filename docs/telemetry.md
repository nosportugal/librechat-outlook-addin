# Telemetry

This add-in emits a small set of OpenTelemetry signals for adoption and usage diagnostics. Telemetry is strictly PII-free: it never includes email contents, recipients, subjects, prompts, or identity values.

## What is emitted

- `addin_installs_total` — first successful load, once per browser profile
- `addin_active_days_total` — one event per browser profile per UTC day
- `outlook_actions_total` — starts for summarize, reply, and tone
- `outlook_action_outcomes_total` — completed, failed, or cancelled
- `outlook_action_duration_ms` — total action time
- `outlook_action_phase_duration_ms` — diagnostic breakdown by phase

## Event labels

All labels are bounded to a small set of values so cardinality stays under control:

- `action`: `summarize`, `reply`, `tone`
- `outcome`: `completed`, `failed`, `cancelled`
- `phase`: `reading_email`, `thinking`, `inserting_result`
- `error_kind`: `auth`, `network`, `timeout`, `http_4xx`, `http_5xx`, `office`, `invalid_response`, `unknown`

## Telemetry boundary

- Browser code posts small JSON events to `/__telemetry`.
- The server is responsible for export wiring when `OTEL_EXPORTER_OTLP_ENDPOINT` is configured.
- When the collector endpoint is unset, telemetry is a safe no-op.
- Telemetry failures must never block the add-in.

## Notes

- Adoption telemetry is only recorded after sign-in.
- Install tracking is based on first observed authenticated load, not Outlook's internal install lifecycle.
- Metrics are intentionally small and bounded so they can be compared across deployments without creating new time series.
