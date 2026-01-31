# AppSource Deployment Checklist

## ✅ Completed Items

### Manifest Updates
- [x] Changed ProviderName from "Contoso" to "Chris Aragao"
- [x] Updated Description to: "Using Arquero, Vega-Lite, and Mermaid to power analysts with tools for data processing, visualizations, and process mapping."
- [x] Updated SupportUrl to GitHub Issues: `https://github.com/cbaragao/Juizo/issues`
- [x] Updated AppDomain to `https://cbaragao.github.io`
- [x] Added Privacy Policy URL: `https://cbaragao.github.io/Juizo/privacy.html`
- [x] Added Terms of Use URL: `https://cbaragao.github.io/Juizo/terms.html`
- [x] Manifest validation passed ✅

### Legal Documents Created
- [x] Privacy Policy created at [privacy.html](privacy.html)
  - Explains no data collection
  - Documents code execution security
  - Lists third-party libraries
- [x] Terms of Use created at [terms.html](terms.html)
  - Security warnings about code execution
  - Disclaimer of warranties
  - Limitation of liability
  - License information

### Code Updates
- [x] Updated webpack.config.js production URL to `https://cbaragao.github.io/Juizo/`
- [x] Added security disclaimer to Arquero Query Editor UI
- [x] Performance optimizations implemented (columnar data loading, single sync)
- [x] Auto-update feature for Vega charts implemented

### Documentation
- [x] README.md with proper attributions for Arquero, Vega-Lite, and Mermaid
- [x] Academic citation for Vega-Lite included
- [x] License information documented

## 📋 Pre-Submission Tasks

### 1. Build and Deploy
```bash
# Build production version
npm run build

# Test the production build locally
# Verify all features work with production build

# Commit and push to GitHub
git add .
git commit -m "Production build for AppSource submission"
git push origin main

# Verify GitHub Pages deployment
# Visit: https://cbaragao.github.io/Juizo/taskpane.html
```

### 2. Asset Verification
- [ ] Verify all icons are accessible:
  - https://cbaragao.github.io/Juizo/assets/juizo16.png
  - https://cbaragao.github.io/Juizo/assets/juizo32.png
  - https://cbaragao.github.io/Juizo/assets/juizo64.png
  - https://cbaragao.github.io/Juizo/assets/juizo80.png
- [ ] Verify privacy.html loads correctly
- [ ] Verify terms.html loads correctly
- [ ] Verify taskpane.html loads correctly

### 3. Cross-Platform Testing
Test on:
- [ ] Excel 2016+ on Windows
- [ ] Excel on Windows (Microsoft 365)
- [ ] Excel on Mac (Microsoft 365)
- [ ] Excel on the web
- [ ] Excel on iPad (optional but recommended)

Test scenarios:
- [ ] Create table with sample data
- [ ] Run Arquero query (test code execution)
- [ ] Create Vega-Lite chart
- [ ] Verify chart auto-updates when data changes
- [ ] Create Mermaid diagram
- [ ] Export results to Excel
- [ ] Test error handling
- [ ] Test with large datasets (performance)

### 4. Microsoft Partner Center Account
- [ ] Create Microsoft Partner Center account at https://partner.microsoft.com/dashboard
- [ ] Complete developer profile
- [ ] Set up payout and tax information (if selling)

### 5. Store Listing Assets
Prepare the following:

#### Screenshots (Required)
- [ ] At least 1 screenshot (1366 x 768 pixels recommended)
- [ ] Show key features: Arquero editor, Vega chart, Mermaid diagram
- [ ] Maximum 5 screenshots

#### Video (Optional but Recommended)
- [ ] Demo video showing:
  - Loading data from Excel table
  - Running Arquero transformation
  - Creating visualization
  - Auto-update feature

#### Marketing Copy
- [ ] Short description (80 characters max)
  - Suggestion: "Transform, visualize, and diagram your Excel data with Arquero, Vega-Lite, and Mermaid"
  
- [ ] Long description (4000 characters max)
  - Expand on features
  - Benefits
  - Use cases
  - Link to documentation

- [ ] Search keywords (5-7 recommended)
  - Suggestion: "data transformation", "visualization", "Arquero", "Vega-Lite", "Mermaid", "analytics", "diagrams"

#### Icon for Store
- [ ] 128x128 PNG icon for store listing

### 6. Security Review Preparation

Microsoft will review the code execution feature. Be prepared to explain:
- [ ] Users execute their own code on their own data
- [ ] Code runs client-side only (sandboxed in browser)
- [ ] Clear security warnings in UI (✅ already added)
- [ ] Documented in Privacy Policy and Terms of Use (✅ done)
- [ ] No external code injection - users write code themselves
- [ ] Legitimate use case for data analysts

### 7. Testing Documentation
Create test notes including:
- [ ] Test cases passed
- [ ] Known limitations (if any)
- [ ] Browser compatibility
- [ ] Performance benchmarks

### 8. Final Validation
- [ ] Run `npm run validate` one more time
- [ ] Check all URLs are HTTPS
- [ ] Verify GitHub Pages is enabled and working
- [ ] Test sideloading the production build
- [ ] Review all error messages for professionalism
- [ ] Check console for errors/warnings

## 🚀 Submission Steps

1. **Go to Partner Center Dashboard**
   - https://partner.microsoft.com/dashboard/marketplace-offers/overview

2. **Create New Offer**
   - Select "Office Add-in"
   - Choose "Excel"

3. **Upload Manifest**
   - Upload your validated manifest.xml

4. **Complete Store Listing**
   - Add screenshots
   - Add description
   - Add support information
   - Add privacy policy URL
   - Add terms of use URL

5. **Submit for Certification**
   - Review submission
   - Submit

6. **Certification Process**
   - Typically takes 1-5 business days
   - Microsoft will test functionality
   - Security review for code execution
   - May request changes

7. **Go Live**
   - Once approved, publish to AppSource
   - Monitor user feedback
   - Respond to issues promptly

## 📝 Notes

### Security Consideration
The `new Function()` code execution in Arquero editor is a legitimate feature but may raise questions. Key points to emphasize:
- It's a **feature, not a bug** - allows data analysts to write transformations
- Users only execute code they write themselves
- All execution is client-side (sandboxed)
- Similar to Excel formulas or Power Query M code
- Clear warnings provided to users
- Documented in terms of use

### Support Strategy
- GitHub Issues for bug reports
- Consider adding:
  - Examples repository
  - Video tutorials
  - Sample workbooks

### Future Considerations
- Add more example queries in UI
- Create a library of common transformations
- Add query validation/linting
- Consider telemetry (with user consent) for improving features

## 🔗 Important Links

- **Partner Center:** https://partner.microsoft.com/dashboard
- **AppSource:** https://appsource.microsoft.com/
- **Validation Guide:** https://learn.microsoft.com/office/dev/add-ins/testing/troubleshoot-manifest
- **Submission Guide:** https://learn.microsoft.com/office/dev/store/submit-to-appsource-via-partner-center

---

**Status:** Ready for pre-submission testing ✅  
**Next Step:** Build production version and deploy to GitHub Pages
