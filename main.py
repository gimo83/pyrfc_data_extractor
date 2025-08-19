import os
from configparser import ConfigParser
import pandas as pd


po_list = [
    '0004056101',
    '0004055874',
    '0004055873',
    '0004057015',
    '4500019671'
]


def config_rfcsdk():
    proj_dir = os.getcwd()
    os.environ['SAPNWRFC_HOME'] = proj_dir + '\\nwrfcsdk' 



def load_config():
    configur = ConfigParser()
    configur.read('config.ini')
    return configur


def main():
    config_rfcsdk()

    from pyrfc_extractor import Connection

    config = load_config()

    conn = Connection(
        ashost=config.get('SAP','host'), 
        sysnr=config.get('SAP','instance'), 
        client=config.get('SAP','client'), 
        user=config.get('SAP','username'), 
        passwd=config.get('SAP','password'), 
        lang="EN", 
    )

    po_data = []
    for po in po_list:
        result = conn.get_po_invoices(po)
        po_data += result
    
    if po_data:
        df_po_ri = pd.DataFrame(data=po_data)
        df_po_ri['AWKEY'] = df_po_ri['BELNR'] + df_po_ri['GJAHR']
        AWKEY_docs = list(df_po_ri['AWKEY'].drop_duplicates())
    
    fi_doc_hd = []
    for objKey in AWKEY_docs:
        result = conn.get_fi_doc(objKey)
        fi_doc_hd += result
    
    df_fi_doc_hd = pd.DataFrame(data=fi_doc_hd)

    fi_doc_it = []
    for fi_doc in fi_doc_hd:
        result = conn.get_fi_doc_items(fi_doc['BUKRS'],fi_doc['BELNR'],fi_doc['GJAHR'])
        fi_doc_it += result
    
    if fi_doc_it:
        df_fi_doc_it = pd.DataFrame(data=fi_doc_it)
        df_fi_doc_it.loc[ df_fi_doc_it['AUGDT'] == '00000000' ,'AUGDT'] = ''
        df_fi_doc_it.loc[ df_fi_doc_it['AUGCP'] == '00000000' ,'AUGCP'] = ''
    
    if df_fi_doc_hd is not None and df_fi_doc_it is not None and df_po_ri is not None:
        df_fi_doc = pd.merge(df_fi_doc_hd,df_fi_doc_it,on=['BUKRS','BELNR','GJAHR'])
        df_output = pd.merge(df_po_ri,df_fi_doc,on='AWKEY')
        with pd.ExcelWriter("output_report\\po_invoice_data.xlsx") as writer:
            df_summary = df_output[['EBELN','GJAHR_x','BELNR_x','BELNR_y','GJAHR_y','AUGDT','AUGCP','AUGBL']]
            df_summary.to_excel(writer,sheet_name="summary")
            df_po_ri.to_excel(writer,sheet_name="po_invoices")
            df_fi_doc_hd.to_excel(writer,sheet_name="fi_doc_header")
            df_fi_doc_it.to_excel(writer,sheet_name="fi_doc_item")



if __name__ == '__main__':
    main()