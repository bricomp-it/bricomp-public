# Add-BCMailboxDomainAlias

Adds a domain-based SMTP alias to Exchange recipients that have automatic
email address policy management turned off. The alias is built from each
recipient's own Exchange Alias attribute, so every affected user gets a
unique address (`jsmith@your-domain.com`) rather than one address being
reused across multiple mailboxes.

Both recipient sources are checked by default:

- **On-premises mailboxes**, via `Get-Mailbox` / `Set-Mailbox`.
- **Hybrid remote mailboxes** -- on-premises AD objects whose actual
  mailbox is hosted in Exchange Online -- via `Get-RemoteMailbox` /
  `Set-RemoteMailbox`.

If either cmdlet isn't available in the current session, that source is
skipped with a warning instead of failing the run. The script only stops
if neither is available. In a hybrid environment, most real users are
typically remote mailboxes rather than on-premises mailboxes, so both
sources need to be checked to get a complete picture.

## Why this exists

When a recipient has `EmailAddressPolicyEnabled` set to `$false`, its SMTP
addresses are no longer kept in sync automatically when an email address
policy is applied or updated. If you need to roll out a new domain to
exactly that subset of recipients, you have to add the address manually,
one at a time. This script automates that, safely and repeatably, across
both on-premises and hybrid recipients.

## Requirements

- Must be run from an Exchange Management Shell session. Tested against
  on-premises Exchange Server, including Exchange Server SE. At least one
  of `Get-Mailbox`/`Set-Mailbox` or `Get-RemoteMailbox`/`Set-RemoteMailbox`
  must be available; the other is skipped with a warning if missing.
- Permissions sufficient to run `Set-Mailbox` and/or `Set-RemoteMailbox`
  against the target recipients.
- This is a runtime prerequisite only. No Microsoft code is bundled with
  or redistributed by this script.

## A note on Get-Mailbox vs. Get-RemoteMailbox

In a hybrid Exchange deployment, a user's real mailbox can live in
Exchange Online while the on-premises Active Directory only holds a
"remote mailbox" stub object for them. `Get-Mailbox` only sees mailboxes
actually hosted on-premises; `Get-RemoteMailbox` only sees those stub
objects. If you run this from Exchange Management Shell and get
unexpectedly few (or zero) results, check which type your target users
actually are:

```powershell
Get-Recipient -Identity "someone" | Format-List Name, RecipientTypeDetails
```

`RemoteUserMailbox` means that user is a hybrid remote mailbox, not an
on-premises one -- this script already checks both by default, but it's
useful to know which source is producing your results.

## A note on Identity values

The script identifies each recipient to `Set-Mailbox`/`Set-RemoteMailbox`
by GUID rather than by name or alias. Exchange Management Shell connects
through implicit remoting even for "local" sessions, so objects returned
by `Get-Mailbox`/`Get-RemoteMailbox` are deserialized copies. Passing a
deserialized object's `Identity` property straight into another cmdlet's
`-Identity` parameter can fail with an error like:

```
Cannot process argument transformation on parameter 'Identity'. Cannot
convert the "..." value of type "Deserialized.Microsoft.Exchange.Data.
Directory.ADObjectId" to type "Microsoft.Exchange.Configuration.Tasks.
RemoteMailboxIdParameter".
```

GUID is a plain string every Exchange recipient cmdlet accepts as an
identity, so it avoids this failure entirely.

## Usage

Review changes first without modifying anything:

```powershell
.\Add-BCMailboxDomainAlias.ps1 -Domain "alias.contoso.com" -WhatIf
```

Apply the changes:

```powershell
.\Add-BCMailboxDomainAlias.ps1 -Domain "alias.contoso.com"
```

Apply the changes and capture full per-mailbox results:

```powershell
.\Add-BCMailboxDomainAlias.ps1 -Domain "alias.contoso.com" -PassThru |
    Export-Csv -Path "C:\Temp\AliasResults.csv" -NoTypeInformation
```

## Parameters

| Parameter                     | Required | Default              | Description |
|--------------------------------|----------|-----------------------|-------------|
| `-Domain`                      | Yes      | -                     | Domain appended to each recipient's Alias to build the new address. |
| `-RecipientTypeDetails`        | No       | `UserMailbox`         | On-premises mailbox type(s) to evaluate (`Get-Mailbox -RecipientTypeDetails`). |
| `-RemoteRecipientTypeDetails`  | No       | `RemoteUserMailbox`   | Hybrid remote mailbox type(s) to evaluate. `Get-RemoteMailbox` has no built-in type filter, so this is applied client-side. Other values: `RemoteRoomMailbox`, `RemoteEquipmentMailbox`, `RemoteSharedMailbox`, `RemoteTeamMailbox`. |
| `-LogPath`                     | No       | Timestamped file in the current user's Documents folder | Transcript log location. |
| `-PassThru`                    | No       | Off                   | Also emit per-recipient result objects to the pipeline (includes a `Source` column: `Mailbox` or `RemoteMailbox`). |

## Idempotency

The script checks each recipient's existing proxy addresses (case-insensitive)
before making a change. If the target address is already present, that
recipient is logged as `NoChangeNeeded` and skipped. Running the script
repeatedly against the same environment is safe -- on the second and later
runs, every previously-updated recipient reports `NoChangeNeeded`.

## Output

A summary is printed at the end of every run:

```
===== Summary =====
Attempted:        42 (Mailbox: 3, RemoteMailbox: 39)
Changed:          40
Would change:     0
No change needed: 2
Failed:           0
Log:              C:\Users\you\Documents\Add-BCMailboxDomainAlias_20260819_120000.log
```

- **Attempted** -- recipients evaluated that had `EmailAddressPolicyEnabled -eq $false`, broken down by source.
- **Changed** -- the alias address was successfully added.
- **Would change** -- would have been added, but the run was a `-WhatIf` dry run (or the change was declined via `-Confirm`) so nothing was actually modified.
- **No change needed** -- the recipient already had the target address.
- **Failed** -- the add operation threw an error (logged with the specific error message, both to the console and the transcript log).

## License

MIT. See `LICENSE` in this folder.
