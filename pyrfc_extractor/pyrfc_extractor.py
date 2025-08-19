import logging
from collections.abc import Iterable

import pyrfc

from .utils import conver_results_to_object

logger = logging.getLogger(__name__)

# Protect against missing SAP NW SDK prerequisite for pyrfc
try:
    BaseConnection = pyrfc.Connection
    from pyrfc import ABAPRuntimeError
except AttributeError:
    logger.error("pyrfc not configured correctly, likely due to missing SAP NW SDK.")
    BaseConnection = object


class Connection(BaseConnection):
    def get_po_invoices(self,po):
        client = self.get_connection_attributes()['client']
        po_results = []
        kwargs = dict(
                    QUERY_TABLE = 'EKBE',
                    DELIMITER = '|',
                    NO_DATA = '',
                    FIELDS = [ 
                        {'FIELDNAME':'EBELN'} ,
                        {'FIELDNAME':'EBELP'} ,
                        {'FIELDNAME':'ZEKKN'} ,
                        {'FIELDNAME':'VGABE'} ,
                        {'FIELDNAME':'GJAHR'} ,
                        {'FIELDNAME':'BELNR'} ,
                        {'FIELDNAME':'BUZEI'} ,
                            ],
                    ROWSKIPS = 0 ,
                    ROWCOUNT = 500,
                    OPTIONS = [
                        "MANDT = '"+ client +"'", 
                        " AND EBELN = '"+ po +"'",
                        " AND VGABE = '2'"    # received invoice transaction only
                        ]
                )
        try:
            result = self.call('RFC_READ_TABLE',**kwargs)
            po_results = conver_results_to_object(result)
        except ABAPRuntimeError as error:
            print(error.message)
        
        return po_results
    
    def get_fi_doc(self,ObjectKey):
        po_results = []
        client = self.get_connection_attributes()['client']

        kwargs = dict(
                    QUERY_TABLE = 'BKPF',
                    DELIMITER = '|',
                    NO_DATA = '',
                    FIELDS = [ 
                        {'FIELDNAME':'BUKRS'} ,
                        {'FIELDNAME':'BELNR'} ,
                        {'FIELDNAME':'GJAHR'} ,
                        {'FIELDNAME':'BLART'} ,
                        {'FIELDNAME':'AWTYP'} ,
                        {'FIELDNAME':'AWKEY'} ,
                            ],
                    ROWSKIPS = 0 ,
                    ROWCOUNT = 500,
                    OPTIONS = [
                        "MANDT = '"+ client +"'", 
                        " AND AWKEY = '"+ ObjectKey +"'",
                        ]
                )
        try:
            result = self.call('RFC_READ_TABLE',**kwargs)
            po_results = conver_results_to_object(result)
        except ABAPRuntimeError as error:
            print(error.message)
        
        return po_results
    
    def get_fi_doc_items(self,BUKRS,BELNR,GJAHR):
        results = []
        client = self.get_connection_attributes()['client']
        kwargs = dict(
                    QUERY_TABLE = 'BSEG',
                    DELIMITER = '|',
                    NO_DATA = '',
                    FIELDS = [ 
                        {'FIELDNAME':'BUKRS'} ,
                        {'FIELDNAME':'BELNR'} ,
                        {'FIELDNAME':'GJAHR'} ,
                        {'FIELDNAME':'BUZEI'} ,
                        {'FIELDNAME':'AUGDT'} ,
                        {'FIELDNAME':'AUGCP'} ,
                        {'FIELDNAME':'AUGBL'} ,
                        {'FIELDNAME':'BSCHL'} ,
                            ],
                    ROWSKIPS = 0 ,
                    ROWCOUNT = 500,
                    OPTIONS = [
                        "MANDT = '"+ client +"'", 
                        " AND BUKRS = '"+ BUKRS +"'",
                        " AND BELNR = '"+ BELNR +"'",
                        " AND GJAHR = '"+ GJAHR +"'",
                        " AND BSCHL = '31'",             # invoice line item only
                        ]
                )
        try:
            result = self.call('RFC_READ_TABLE',**kwargs)
            results = conver_results_to_object(result)
        except ABAPRuntimeError as error:
            print(error.message)
        
        return results
