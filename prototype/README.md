# Practice accounts MVP

This standalone app lives on `prototype/multi-practice-auth`, at `prototype/`.
It reuses the original dashboard calculations and controls with a separate Supabase
adapter. It does not change GitHub Pages, connect to Microsoft/OneDrive, or read the
live site's data. The original root dashboard sources remain unchanged.

## What is working

- Email/password sign-in, invited account creation and password recovery through Supabase Auth.
- Platform administrator: create practices and invite practice administrators.
- Practice administrator: invite viewers, approve and suspend their practice's members.
- Viewer: read only the practices for which they have approved membership.
- Direct CSV uploads: validate a single tenant, group requests by minute, and append
  only minutes later than that practice's server-owned cutoff. Import and cutoff
  update occur in one locked PostgreSQL transaction.
- Rapid Health monthly weekday averages, including zero days and partial coverage.
- Activity IQ historical activity, waiting ranges, median waits, booking times and Signposting.
- Capacity IQ urgent/routine capacity and Day View.
- Forecast IQ urgent predictions/analytics, routine forecasts and routine capacity/waiting times.
- Workforce Need vs appointments; personal, slot, people and practice settings.
- Slot/Signposting direct imports and existing version-9 settings JSON imports.
- Per-practice data freshness, import history and an audit log.
- A separate, clearly labelled in-memory demo with fictional practices and role previews.

Rapid Health retains daily Other enquiry counts, total request counts and coverage.
Slot Reports retain the existing eight reporting aggregates (including clinician names
needed for People/Workforce); Signposting retains date/time/destination only. They are
compressed before saving. Original CSVs and patient identifiers are not retained.
Each Slot/Signposting source is replaced atomically, with revision checks to reject
stale concurrent writes. Same filenames match case-insensitively. Renamed refreshed
exports must select their existing source; use new sources only for non-overlapping
periods. Imports currently support one non-empty CSV/JSON up to 50 MB, with a 12 MB
compressed snapshot limit. Keep original exports outside this prototype.

Shared settings are stored per practice with optimistic concurrency checks. Browser
preferences are scoped by both user and practice before the dashboard initializes.
The dashboard frame receives a practice-bound bridge and never handles an API secret
or Microsoft token. Every bridge read rechecks approved access; database RLS and
checked RPCs enforce access independently of the UI. Revocation closes the dashboard
on the next workspace check. Anonymous direct dashboard URLs display no practice data.
Apply `supabase/migrations/20261105182856_practice_analytics.sql` after the initial
migration. Additional checks: `node tests/analytics-database.mjs` with PGLITE_MODULE.
The original Access IQ coming-soon placeholder remains a placeholder.

## Connect an isolated free test project

1. Create a **new test** Supabase project. Do not use a production project.
2. Inspect its tables, then apply `supabase/migrations/202611050001_practice_mvp.sql`.
3. In Supabase Authentication, enable email/password and email confirmation. Set the
   Site URL and allowed redirect URLs to this preview's exact origin/path.
   Configure an SMTP service before inviting external test users. Supabase's
   default email sender is limited to project-team addresses.
4. Configure the designated owner's verified email in `iq_private.platform_bootstrap`
   using an administrative connection, then open the preview at `?setup=1`.
   The owner creates their own password and verifies their email. The checked
   `iq_claim_platform_admin` RPC grants the first platform role only to that
   exact verified email and consumes the setup record. A URL flag alone grants
   no access. Reopening setup cannot reinstate a revoked administrator.

5. Set `supabaseUrl` and `publishableKey` in `prototype/config.js`. These two values
   are public client configuration. Never put a service-role/secret key in this app.
6. Build/redeploy the private preview. Sign in, create two fictional practices with
   different tenant names, invite their admins/users and test cross-practice access.

Invitation emails are submitted by the `iq-send-invite` Supabase Edge Function through
Brevo. Deploy the function with its own `getUser()` authentication and `verify_jwt=false`
(the function verifies the session itself, including projects with newer signing keys).
Set `BREVO_API_KEY` and `INVITE_FROM_EMAIL` as Supabase Edge Function secrets. The sender
must be verified in Brevo; never put the API key in browser configuration. The function
uses the caller's checked invitation RPC, without a service-role key. Its destination
URL and browser origin are fixed to this prototype. Changing hosting needs updating them.
The UI confirms submission to the provider, not delivery to an inbox. Check Brevo's
transactional logs for bounces and delivery. No recipient contacts list is created.
Account-verification and password-reset emails separately require Supabase custom SMTP
for external users; Brevo SMTP credentials can be configured in Supabase Auth. Set the
Auth Site URL and redirect allowlist too. Hosting currently remains owner-private;
external users also need hosting access enabled before they can open invitations.
New members remain pending after accepting.
Only the platform administrator can appoint or change practice administrators.
Suspension is checked in the database on subsequent requests, including existing sessions.

## Run and test

Serve the repository over HTTP and open `/prototype/`. Do not open modules via `file:`.
The demo does not use a backend or store anything between visits. Demo mode is for
UI exploration, not evidence of authentication security.

```text
node --test tests/*.test.cjs
PGLITE_MODULE=/absolute/path/to/pglite/dist/index.js node tests/prototype-database.mjs
node prototype/build.mjs /absolute/path/to/separate-preview/dist
```

The database test runs the actual migration in PostgreSQL via PGlite, with a minimal
Supabase Auth schema stub. It exercises RLS and RPC permissions using non-owner roles.
It does not test Supabase email delivery, hosted Auth, or browser sessions; those checks
require the connected test project and real test identities.

Security definitions: all client writes use checked RPCs; direct table writes and
anonymous reads are denied. Membership approval is never taken from browser metadata.
Each RPC pins its search path and checks the authenticated caller. RLS remains in
place even when someone changes browser controls or calls database URLs directly.

References: https://supabase.com/docs/guides/database/postgres/row-level-security
and https://supabase.com/docs/guides/database/functions
