# %%
import os
import subprocess
import sys

# %% [markdown]
# # Configure NetWeaver SDK library location

# %%
proj_dir = os.getcwd()

# %%
os.environ['SAPNWRFC_HOME'] = proj_dir + '\\nwrfcsdk' 

# %%
#os.environ['SAPNWRFC_HOME'] = proj_dir + '\\nwrfc750\\nwrfcsdk' 

# %%
#os.environ['PATH'] = os.environ['PATH'] + ';' + proj_dir + '\\nwrfc750\\nwrfcsdk\\lib'

# %%
#os.environ['PATH'] = os.environ['PATH'] + ';' + proj_dir + '\\nwrfc750\\nwrfcsdk\\bin'

# %%
#pip install pyrfc-read

# %% [markdown]
# # 1. Locad configuration

# %%
from configparser import ConfigParser

# %%
configur = ConfigParser()

# %%
configur.read('config.ini')

# %%
ASHOST=configur.get('SAP','host')
CLIENT=configur.get('SAP','client')
SYSNR=configur.get('SAP','instance')
USER=configur.get('SAP','username')
PASSWD=configur.get('SAP','password')

# %% [markdown]
# ## Import NetWeaver RFC libriary in Python

# %%
try:
    from pyrfc import ABAPRuntimeError
except ImportError:
    # Install pyrfc using pip
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyrfc==3.3.1"])
    from pyrfc import ABAPRuntimeError

# %%
try:
    from pyrfc_read import Connection
except:
    # Install pyrfc using pip
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyrfc-read"])
    from pyrfc_read import Connection


# %% [markdown]
# ## Open connection to SAP System

# %%
conn = Connection(
    ashost=ASHOST, 
    sysnr=SYSNR, 
    client=CLIENT, 
    user=USER, 
    passwd=PASSWD,
    lang="EN"
    )

# %% [markdown]
# ## Call SAP RFC 'RFC_PING'

# %%
try:
    result = conn.ping()
    print(result)
except ABAPRuntimeError as error:
    print(error.message)

# %% [markdown]
# # Read PO Detail

# %%
try:
    kwargs = dict(
            PURCHASEORDER='6510090283',
            ITEMS='X',
            ACCOUNT_ASSIGNMENT='X',
            SCHEDULES='X',
            HISTORY='X',
            ITEM_TEXTS='X',
            HEADER_TEXTS='X',
            SERVICES='X',
            CONFIRMATIONS='X',
            SERVICE_TEXTS='X',
            EXTENSIONS='X',
        )
    result = conn.call('BAPI_PO_GETDETAIL',kwargs)
    print(result)
except ABAPRuntimeError as error:
    print(error.message)
else:
    print("un-known error")