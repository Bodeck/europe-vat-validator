# Europe VAT Validator
Excel VBA tool to validate VAT Registration Numbers for EU and UK business customers, using public WebAPIs.
## Usage instructions
1. Open `EuropeVatValidator.xlsm` file.
1. Go to sheet `Validator` and enter/paste VAT numbers to be validated in column `VAT Registration Number`. Valid format is `{ISO Country Code}{vat number}`
1. Select icon `Run Validator` icon from Excel Quick Access Toolbar
1. If you are behind proxy, select option `Connect via proxy server` and enter proxy user and password. Other proxy settings going to be loaded automatically in background.
1. Click `Start` to run validation process or `Cancel` to close user form
1. Once validation is completed message box will pop-up.  

## Notes
* Tested on Windows 10 with Office 2013 (32 bit) and Office 2016 (64 bit)
## Web APIs
### EU Countries

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

As brexit became a fact, HMRC (HM Revenue & Customs) provides separate web service to check UK companies vat numbers.
See [Check a UK VAT number API](https://developer.service.hmrc.gov.uk/api-documentation/docs/api/service/vat-registered-companies-api/1.0) documentation for more details.

## Dependecies

For HTTP communication I have used VBA library availble at [VBA-Web](https://github.com/VBA-tools/VBA-Web) repository.

Where early binding have been used, following references must be set up within VBE:
* MSXML v3.0 (DOMDocument)
* Microsoft Scripting Runtime (Dictionary)

## TODO
- [ ] implement exception handler