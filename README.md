before running, install the necessary packages in requirements.txt by running
<code> 
pip install --upgrade pip
pip install -r requirements.txt 
<\code>
if running into SSL Error and the error seems to asking for the certificate bundle, try to locate the certificate at root and repoint the environment variable to the correct .cert file or try install pip-system-certs first using
<code>
pip install pip-system-certs
<\code>
