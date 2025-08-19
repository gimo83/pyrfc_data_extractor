def conver_results_to_object(result):
    results = []
    for row in result['DATA']:
        
            d = row['WA'].split('|')
            fields = result['FIELDS']
            o = {}
            for i,v in enumerate(d):
                o[fields[i]['FIELDNAME']] = v
            results.append(o)
    return results