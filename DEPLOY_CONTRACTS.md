# Full Contract Template Deployment

Deploy these files to the existing Google Apps Script project:

- `Code.gs`
- `New.gs`

The frontend `index.html` already uses this deployment URL:

`https://script.google.com/macros/s/AKfycbynaI2-iS40nJhK3hZSw6LWTJ1VpD60SkkogU_4REvINHl-jVZbC9Zt4V5t5oLYeqna/exec`

## Deployment steps

1. Replace the Apps Script project's `Code.gs` and `New.gs` contents.
2. Select **Deploy > Manage deployments**.
3. Edit the existing web-app deployment.
4. Choose **New version**.
5. Execute as **Me**.
6. Keep access compatible with the existing dashboard.
7. Deploy.

## Verify before generating

Open:

`https://script.google.com/macros/s/AKfycbynaI2-iS40nJhK3hZSw6LWTJ1VpD60SkkogU_4REvINHl-jVZbC9Zt4V5t5oLYeqna/exec?action=testDocuments`

Expected result:

- `ok: true`
- Contract paragraph count: at least `117`
- Extension paragraph count: at least `27`

The backend refuses to generate a document if any required section is missing
or if any `{{VARIABLE}}` remains unfilled.

## Verified templates

- Full lease contract:
  `https://docs.google.com/document/d/1ZjHj8rmzSyJkcs6hpOOvCLev8c_fFOl3vn-cYSt7VWU/edit`
- Full extension agreement:
  `https://docs.google.com/document/d/1_5InwAOi9cXo1sccip0jJsvcHn-CYJfFFURhh0Ht_hE/edit`

Local exact round-trip samples:

- `sample_generated_contract.docx`
- `sample_generated_extension.docx`
