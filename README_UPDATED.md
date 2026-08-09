# Updated USPTO Trademark Scraper

This version preserves the existing React/Vite frontend and Google Apps Script workflow while adding:

- International Class
- Goods / Services
- Application Filing Date
- Status Date
- Owner Name(s)
- Owner Address(es)
- Legal Entity Type
- State/Country Where Organized
- Mark Type
- Register
- Basis
- TM5 Common Status Descriptor
- Mark Description
- Disclaimer
- Existing attorney detection
- Existing correspondent contact fields
- Detailed CSV export

## Deployment

### Google Apps Script
Use `google-apps-script/Code.gs` in the existing Google Apps Script project.

After saving, deploy a new version of the existing Web App. Keep the same Web App URL if possible.

### Vercel
Commit/push the repository to the GitHub repository connected to Vercel. Vercel will redeploy automatically.

### Important
Do not put API keys in the React frontend. The current scraper does not require the `USPTO_API_KEY` variable.

Test first with serial `50038347`.
