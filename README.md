# n8n-nodes-caldav-interface

CalDAV community nodes for [n8n](https://n8n.io/). This package lets workflows and AI agents create calendars, manage events, query schedules, and compute free/busy availability against standards-based CalDAV servers such as Radicale.

## Installation

Install the package from the n8n community nodes UI or with npm in a self-hosted n8n setup:

```bash
npm install n8n-nodes-caldav-interface
```

Then restart n8n and add the `CalDAV Interface` node to a workflow.

## Operations

The `CalDAV Interface` node is programmatic and supports these resources and operations:

- `Calendar`
  - `List`
  - `Create`
  - `Delete`
  - `Get`
- `Event`
  - `Create`
  - `Get` by identifier
  - `Get` by date range
  - `Update`
  - `Delete`
- `Query`
  - `Filter Events`
  - `Get Free/Busy`

The node handles CalDAV-specific WebDAV methods such as `PROPFIND`, `REPORT`, `MKCALENDAR`, and standard `PUT` / `DELETE`, along with XML Multi-Status parsing and iCalendar payload generation.

## Credentials

Create a `CalDAV API` credential with:

- `Base URL`: CalDAV server root or account URL, for example `https://calendar.example.com` or `http://localhost:5232`
- `Username`
- `Password`
- `Calendar Home Path` optional override, useful when discovery is incomplete on a server
- `Default Timezone` used for natural language date handling
- `Allow Unauthorized Certificates` only for trusted self-signed environments

Authentication uses HTTP Basic Auth.

### Credential Validation

The credential test sends an authenticated request to the configured base URL. For stricter setups, point `Base URL` at the account or collection root that the server allows authenticated access to.

## AI Agent Tool Support

The node is marked `usableAsTool: true` and its descriptions are optimized for agent use. Common patterns:

- `Event -> Create`
  - `When`: `next Tuesday at 3pm for 45 minutes`
  - `Title`: `Customer renewal call`
- `Event -> Get` with `Date Range`
  - `Time Range`: `tomorrow`
- `Query -> Get Free/Busy`
  - `Time Range`: `Friday afternoon`
  - `Minimum Duration Minutes`: `120`

Recommended agent flow:

1. Use `Query -> Filter Events` or `Event -> Get` with a date range to find candidate events.
2. Use `Event -> Update` or `Event -> Delete` with the returned `href`, filename, or UID.
3. Use `Query -> Get Free/Busy` before scheduling when conflict detection matters.

## Example Workflows

### Manual workflow

1. Add `CalDAV Interface`.
2. Select `Event -> Create`.
3. Pick a calendar.
4. Set `Title`, `When`, and optional `Location`.
5. Execute the node to create the event.

### AI scheduling workflow

1. Use an AI Agent node.
2. Add `CalDAV Interface` as a tool.
3. Prompt with instructions like:

```text
Schedule a dentist appointment for April 15th at 10am for 1 hour in my Personal calendar.
```

The agent can create the event directly and the node will return a structured payload plus a human-readable summary.

## Radicale Notes

Radicale is a good local target for validation.

- Default URL: `http://localhost:5232`
- Common base URL choices:
  - `http://localhost:5232/`
  - `http://localhost:5232/<user>/`
- If principal discovery is limited, set `Calendar Home Path` explicitly, for example `/<user>/`

Example local credential values:

- `Base URL`: `http://localhost:5232`
- `Username`: your Radicale username
- `Password`: your Radicale password
- `Calendar Home Path`: `/<user>/`

## Troubleshooting

- `401 Unauthorized`
  - Confirm the username/password and that the base URL is inside the authenticated CalDAV area.
- `No calendars found`
  - Set `Calendar Home Path` explicitly. Some servers expose calendars correctly only from a user collection path.
- `Unable to parse time range`
  - Use a more explicit value such as `2026-04-15 09:00 to 2026-04-15 17:00`.
- `412 Precondition Failed`
  - The ETag changed. Re-fetch the event and retry the update or delete with the latest `etag`.
- Self-signed TLS certificate errors
  - Enable `Allow Unauthorized Certificates` only in trusted development environments.

## Compatibility

This package targets current n8n community-node packaging and standards-based CalDAV servers. It was designed for Radicale and generic CalDAV collections that support:

- Basic Auth
- `PROPFIND`
- `REPORT`
- `PUT`
- `DELETE`
- `MKCALENDAR`

Live server verification depends on your CalDAV server configuration. The repository build and lint steps should be run after changes before publishing.

## Resources

- [n8n community nodes documentation](https://docs.n8n.io/integrations/#community-nodes)
- [CalDAV RFC 4791](https://www.rfc-editor.org/rfc/rfc4791)
- [Radicale documentation](https://radicale.org/)

## Version History

### 0.1.0

Initial CalDAV community-node release with calendar management, event management, date-range querying, free/busy calculation, and AI tool support.
