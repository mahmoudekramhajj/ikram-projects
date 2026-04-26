# Deploy Google Apps Script
1. Run `clasp push` to upload code
2. Get existing deployment ID: `clasp deployments`
3. Update existing deployment: `clasp deploy --deploymentId <ID> -d "Update"`
4. NEVER create a new deployment
5. Verify the web app URL is accessible
6. Report the deployment URL to user
