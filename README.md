# SHREE RCH Template API

Vercel serverless API that receives the existing SHREE RCH template request:

- `typData=mother` or `child`
- `baseData=<base64 encoded JSON locationData>`

Endpoint:

`POST /api/template`

The generated XLSX contains exactly three location dropdown columns on `Main`:

1. Health Facility
2. SubCentre
3. Village

SubCentre is dependent on Health Facility, and Village is dependent on Health Facility + SubCentre.

No Excel auto-filter is added.

## Deploy

1. Push this repository to GitHub.
2. Import the repository into Vercel.
3. Deploy.
4. Test:

`GET https://YOUR-DOMAIN.vercel.app/api/template`

Then update the extension's old template URL to:

`https://YOUR-DOMAIN.vercel.app/api/template`
