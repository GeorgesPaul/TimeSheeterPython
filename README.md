# TimeSheeterPython

A web app that generates invoices and timesheets from Google Calendar events.

- **Simple invoice** — fill in hours and a description in the browser, generate a PDF instantly. No Google Calendar needed.
- **Timesheet invoice** — pulls worked hours from Google Calendar events tagged with `@clientname`, generates a timesheet and invoice PDF.

---

## Quick start (simple invoice with dummy data)

### 1. Install dependencies

Requires Python 3.10+. Using [uv](https://github.com/astral-sh/uv) (recommended):

```
uv sync
```

Or with pip:

```
pip install flask flask-wtf google-api-python-client google-auth google-auth-oauthlib \
            pandas pypdf python-dateutil pyyaml tabulate wtforms xhtml2pdf
```

### 2. Set up config files

Copy the example files and rename them:

| Example file | Rename to |
|---|---|
| `example_clients.yaml` | `clients.yaml` |
| `templates/example_invoice_template.html` | `templates/invoice.html` |

```
cp example_clients.yaml clients.yaml
cp templates/example_invoice_template.html templates/invoice.html
```

The example files contain dummy data for two fake clients so you can generate an invoice right away.

### 3. Personalise your invoice template

Open `templates/invoice.html` and replace the hardcoded seller details in the top-right block with your own:

- Name and company/brand name
- Address
- VAT number
- Email
- IBAN, BIC/SWIFT, bank name

### 4. Run the app

```
python app.py
```

Then open [http://localhost:5000](http://localhost:5000) in your browser.

### 5. Generate a simple invoice

1. Set an **Invoice Date**.
2. Under **Simple Invoice**, select a client (e.g. *Acme Corporation B.V.*).
3. Enter an invoice number, hours, and a description.
4. Click **Generate Simple Invoice** — a PDF opens automatically.

---

## Timesheet invoice (Google Calendar)

To generate invoices from your Google Calendar you need two additional files:

### client_secrets.json

Download your OAuth 2.0 credentials from the [Google Cloud Console](https://console.cloud.google.com/):

1. Create a project and enable the **Google Calendar API**.
2. Create an **OAuth 2.0 Client ID** (Desktop app).
3. Download the JSON and save it as `client_secrets.json` in the project root.

The first time you run the app and trigger a timesheet generation, a browser window will open for Google sign-in. This creates a `token.pickle` file so you only need to authenticate once.

### config.ini

Create `config.ini` in the project root:

```ini
[Google Calendar]
CalID = your-calendar-id@group.calendar.google.com
```

Find your calendar ID in Google Calendar → Settings → select your calendar → scroll to *Integrate calendar*.

### Tag your calendar events

Add `@ClientKey` to the description of any Google Calendar event where `ClientKey` matches a key in `clients.yaml` (e.g. `@AcmeCorp`). The app matches tags to clients automatically.

### Generate a timesheet invoice

1. Set **Invoice Date**, **Start Date**, and **End Date**.
2. Select one or more clients.
3. Choose timesheet columns and whether to append the timesheet to the invoice PDF.
4. Click **Generate Timesheet and Invoice**.

---

## clients.yaml structure

Each client entry looks like this:

```yaml
Clients:
  ClientKey:                          # used for @ClientKey tags in calendar events
    alias_tag: Short name
    registration_name: Legal Name B.V.
    trade_name: Trade Name B.V.
    hourly_rate: 123.00
    currency: "&euro;"
    vat_number: NL000000000B01
    chamber_of_commerce_number: 00000000
    invoice_address:
      street_name: Example Street
      house_number: 1
      other_address_details:          # optional, e.g. "Floor 3" — leave blank if unused
      postal_code: 1234 AB
      city: Amsterdam
      country: Netherlands
    contact_person:
      name: Jane Doe
      email: jane@example.com
      phone: +31600000000
    invoice_email_address: finance@example.com
    payment_terms_days: 14            # due date = invoice date + this many days
```

---

## Files that are gitignored

These files contain personal or sensitive data and are not committed to the repository:

| File | Purpose |
|---|---|
| `clients.yaml` | Your real client data |
| `templates/invoice.html` | Your invoice template with personal/banking details |
| `config.ini` | Google Calendar ID |
| `client_secrets.json` | Google OAuth credentials |
| `token.pickle` | Google auth token (generated on first run) |

Use the `example_*` files as starting points.
