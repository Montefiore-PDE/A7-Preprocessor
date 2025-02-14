Install the necessary packages in requirements.txt by running
<code>
pip install --upgrade pip
pip install -r requirements.txt
</code>

If running into SSL Error and the error seems to asking for the certificate bundle, try to locate the certificate at root and repoint the environment variable to the correct .cert file or try installing pip-system-certs first using
<code>pip install pip-system-certs</code>

At first run, it will create folders called 'SAHRED_DATA' and we need to download fresh copies of Infor and CCX report to support the contract/contract line search <br>
**from Infor:**
* ContractLine.csv
* ContractLineImport.csv
* Manufacturers.csv
* Suppliers.csv
* ItemUOM.csv

**from CCX:**
* ContractOrganization.xlsx (when export, we only need active contract linked to INFOR account)

under this folder, we also will need to keep a fixture file which helps to translate the UOM from multiple sources, this file is uploaded to the repository as well
* UOM.csv

