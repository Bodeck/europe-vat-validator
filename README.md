# Europe VAT Validator
Excel VBA tool to validate VAT Registration Numbers for EU and UK companies. In background it connects public WebAPIs.
## Usage instructions
    1. Open VatValidator.xlsm file 
    2. In sheet "Validator" enter VAT numbers to be validated
    3. Select icon "Run Validator" icon from Excel Quick Access Toolbar
    4. On user form mark option "Connect via proxy server". If your are not behind proxy, skip to point 6.
    5. Enter proxy user and password. Other proxy details should be detected automatically
    6. Click Start to run validation process or Cancel to close user form
    7. Once validation is completed message box will pop-up.  
     
## Web APIs
### EU Countries
![Europe Commition Logo](https://ec.europa.eu/taxation_customs/vies/images/template-2012/logo/logo_en.gif)
Eropean Commission's VIES (VAT Information Electronic System) offers SOAP service which allows to validate VAT Registration Number for EU countries (see table below).
For more details about service see [VIES FAQ](https://ec.europa.eu/taxation_customs/vies/faq.html).
|||||
|--|--|--|--|
|AT-Austria|EE-Estonia|IE-Ireland|PL-Poland
|BE-Belgium|EL-Greece|IT-Italy|PT-Portugal
|BG-Bulgaria|ES-Spain|LT-Lithuania|RO-Romania
|CY-Cyprus|FI-Finland|LU-Luxembourg|SE-Sweden
|CZ-Czech Republic|FR-France|LV-Latvia|SI-Slovenia
|DE-Germany|HR-Croatia|MT-Malta|SK-Slovakia
|DK-Denmark|HU-Hungary|NL-The Netherlands|XI-Northern Ireland

### United Kingdom
![HRMC Logo](https://www.gov.uk/assets/static/gov.uk_logotype_crown_invert_trans-203e1db49d3eff430d7dc450ce723c1002542fe1d2bce661b6d8571f14c1043c.png)
As brexit became a fact, HMRC provides separate web service to check UK companies vat numbers.
See [Check a UK VAT number API](https://developer.service.hmrc.gov.uk/api-documentation/docs/api/service/vat-registered-companies-api/1.0) documentation for more details.

## Dependecies

For HTTP communication I have used VBA library availble at [VBA-Web](https://github.com/VBA-tools/VBA-Web) repository.

Where early binding have been used, following references must be set up within VBE:
* MSXML v3.0 (DOMDocument)
* Microsoft Scripting Runtime (Dictionary)

##TODO
- [ ] implement exception handler