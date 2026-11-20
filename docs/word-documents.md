# Word documents and meeting agendas

Open a Word document in a signed-in Safari SharePoint tab. Moff uses that tab's
same-origin session, not Outlook's Graph permissions. It never exports cookies,
tokens or request digests. With multiple matching SharePoint tabs, select
`--safari-window` and `--safari-tab` after `docs`.

```sh
moff docs read --url 'https://tenant.sharepoint.com/sites/Project/_layouts/15/Doc.aspx?sourcedoc={guid}'
moff outlook calendar agenda --id '<meeting-or-occurrence-id>'
```

`read` returns the main body's paragraphs, heading levels, tables and hyperlink
references. Headers, footers and comments remain in the file but are not extracted.
Deleted tracked text is excluded. Opaque sharing links are resolved through the
signed-in Outlook tab; canonical `Doc.aspx?sourcedoc=...` and `?d=w...` Word links
do not require Graph access. No tab is automatically opened or signed in.

`calendar agenda` searches at most 90 days ahead for the next occurrence of the
same series, in seven-day windows. It returns linked Word document candidates.
A single candidate is selected; multiple links require the caller to choose.
It does not edit the meeting or notify attendees.

## Preview and edit a section

Prepare `edit.json`:

```json
{
  "create_before": "Previous meeting",
  "paragraphs": [
    {"text": "Agenda", "heading_level": 3},
    {"text": "Review the proposed initiatives", "bold": true},
    {"text": "Source PRD", "link": "https://tenant.sharepoint.com/sites/Project/PRD.docx"}
  ]
}
```

```sh
moff docs edit --url '<document-url>' --section 'Next meeting' --body-file edit.json
moff docs --allow-mutation edit --url '<document-url>' --section 'Next meeting' \
  --body-file edit.json --confirm --expected-etag '<etag-from-preview>'
```

Omit `create_before` to append to an existing unique heading. Otherwise it creates
a new section before the named unique heading, at that heading's level. Existing
XML is not reserialized: old paragraphs, formatting, images, comments, bookmarks,
relationships and other package parts are preserved. Editing is append-only.

Apply checks the preview version, then makes one atomic conditional SharePoint
file update. A changed version or active editing lock refuses the write. Close
open editors and wait for their locks to clear rather than forcing an overwrite.
Moff downloads the saved file
and requires an exact byte match. Verification makes at most three read-only
checks, two seconds apart, to allow a just-saved version to become readable.
A failed or uncertain acknowledgement never
triggers a retry: read the document before taking further action.

Limits: 16 MB compressed, 64 MB expanded, 2,048 package parts, and 100 added
paragraphs / 50,000 text bytes. Signed, macro-enabled or protected documents and
documents with track changes enabled cannot be edited.

OneDrive binary transfers also accept an owning `--drive-id` on `get --output`
and `update --file`. File replacement accepts `--expected-etag` to reject changes
since a previously downloaded version. These flags do not grant missing access.
