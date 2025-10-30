/**
 * PicklistExportHelper - Helper class for picklist metadata retrieval
 * Contains ALL functional logic: data retrieval, CSV formatting, file creation
 * This class is completely independent and has no dependencies on the batch class
 */
public class PicklistExportHelper {
    
    private static final String API_VERSION = 'v64.0';
    private static final String CSV_HEADERS = 'Object,Field Label,Field API,Picklist Value Label,Picklist Value API,Status';
    
    // --------------------------------------
    // Inner helper classes
    // --------------------------------------
    public class FieldInfo {
        public String apiName;
        public String label;
    }
    
    public class PicklistValueDetail {
        public String label;
        public String value;
        public Boolean isActive;
    }
    
    public class ProcessingResult {
        public Integer valuesProcessed = 0;
        public Integer inactiveValues = 0;
        public List<String> csvRows = new List<String>();
    }
    
    // --------------------------------------
    // High-level orchestration methods
    // --------------------------------------
    
    /**
     * Process a single object and return CSV rows
     */
    public static ProcessingResult processObject(String objName, String baseUrl, String sessionId) {
        ProcessingResult result = new ProcessingResult();
        
        System.debug('Processing object: ' + objName);
        
        Map<String, FieldInfo> picklistFields = getPicklistFieldsForObject(objName);
        if (picklistFields.isEmpty()) {
            System.debug('  No picklist fields for ' + objName);
            return result;
        }
        
        System.debug('  Found ' + picklistFields.size() + ' picklist fields');
        
        // Resolve EntityDefinition Id once per object
        String entityDefId = resolveEntityDefinitionId(baseUrl, sessionId, objName);
        if (entityDefId != null) {
            System.debug('  EntityDefinition.Id: ' + entityDefId);
        }
        
        for (String fldApi : picklistFields.keySet()) {
            FieldInfo fi = picklistFields.get(fldApi);
            
            List<PicklistValueDetail> values = queryPicklistValuesWithFallback(
                baseUrl, sessionId, objName, entityDefId, fldApi
            );
            
            if (values == null || values.isEmpty()) {
                System.debug('    No values for field: ' + fldApi);
                continue;
            }
            
            System.debug('    Field: ' + fldApi + ' - ' + values.size() + ' values');
            
            for (PicklistValueDetail v : values) {
                String status = (v.isActive == null || v.isActive) ? 'Active' : 'Inactive';
                if (v.isActive != null && !v.isActive) {
                    result.inactiveValues++;
                }
                
                result.csvRows.add(
                    escapeCsv(objName) + ',' +
                    escapeCsv(fi.label) + ',' +
                    escapeCsv(fldApi) + ',' +
                    escapeCsv(v.label) + ',' +
                    escapeCsv(v.value) + ',' +
                    status
                );
                
                result.valuesProcessed++;
            }
        }
        
        System.debug('  Object complete - Values: ' + result.valuesProcessed + ' (Inactive: ' + result.inactiveValues + ')');
        
        return result;
    }
    
    /**
     * Create final CSV file from rows
     */
    public static String createFinalCsvFile(List<String> csvRows) {
        try {
            String csvContent = String.join(csvRows, '\n');
            Blob csvBlob = Blob.valueOf(csvContent);
            
            ContentVersion cv = new ContentVersion();
            cv.Title = 'Picklist_Export_' + Datetime.now().format('yyyyMMdd_HHmmss');
            cv.PathOnClient = cv.Title + '.csv';
            cv.VersionData = csvBlob;
            cv.IsMajorVersion = true;
            cv.ContentLocation = 'S';
            insert cv;
            
            ContentVersion inserted = [SELECT Id, ContentDocumentId FROM ContentVersion WHERE Id = :cv.Id LIMIT 1];
            System.debug('✅ SUCCESS: Final file created - ContentDocumentId: ' + inserted.ContentDocumentId);
            System.debug('File available in Files → Owned by Me');
            
            return inserted.ContentDocumentId;
        } catch (Exception e) {
            System.debug('❌ ERROR creating final file: ' + e.getMessage());
            return null;
        }
    }
    
    /**
     * Create partial CSV file for heap management
     */
    public static void createPartialCsvFile(List<String> csvRows) {
        try {
            if (csvRows == null || csvRows.isEmpty()) return;

            String partName = 'Picklist_Export_Part_' + Datetime.now().getTime();
            String csvContent = String.join(csvRows, '\n');
            Blob csvBlob = Blob.valueOf(csvContent);

            ContentVersion cv = new ContentVersion();
            cv.Title = partName;
            cv.PathOnClient = '/' + partName + '.csv';
            cv.VersionData = csvBlob;
            cv.IsMajorVersion = true;
            cv.ContentLocation = 'S';
            insert cv;

            System.debug('✅ Flushed partial CSV part: ' + partName + ' (rows=' + csvRows.size() + ')');

        } catch (Exception e) {
            System.debug('⚠️ Failed to flush partial CSV: ' + e.getMessage());
        }
    }
    
    /**
     * Get CSV headers
     */
    public static String getCsvHeaders() {
        return CSV_HEADERS;
    }
    
    // --------------------------------------
    // CSV utility methods
    // --------------------------------------
    
    /**
     * Escape CSV values properly
     */
    public static String escapeCsv(String s) {
        if (s == null) return '';
        if (s.contains(',') || s.contains('"') || s.contains('\n') || s.contains('\r')) {
            s = s.replace('"', '""');
            return '"' + s + '"';
        }
        return s;
    }
    
    // --------------------------------------
    // Metadata retrieval methods
    // --------------------------------------
    
    /**
     * Get all picklist fields for an object using Schema describe
     */
    public static Map<String, FieldInfo> getPicklistFieldsForObject(String objectApiName) {
        Map<String, FieldInfo> results = new Map<String, FieldInfo>();
        
        try {
            Schema.SObjectType sType = Schema.getGlobalDescribe().get(objectApiName);
            if (sType == null) {
                System.debug('  Object not found in Schema: ' + objectApiName);
                return results;
            }
            
            Schema.DescribeSObjectResult d = sType.getDescribe();
            Map<String, Schema.SObjectField> fields = d.fields.getMap();
            
            for (String fname : fields.keySet()) {
                Schema.DescribeFieldResult fd = fields.get(fname).getDescribe();
                Schema.DisplayType dt = fd.getType();
                
                if (dt == Schema.DisplayType.PICKLIST || dt == Schema.DisplayType.MULTIPICKLIST) {
                    FieldInfo fi = new FieldInfo();
                    fi.apiName = fd.getName();
                    fi.label = fd.getLabel();
                    results.put(fi.apiName, fi);
                }
            }
        } catch (Exception e) {
            System.debug('ERROR in getPicklistFieldsForObject(' + objectApiName + '): ' + e.getMessage());
        }
        
        return results;
    }
    
    /**
     * Resolve EntityDefinition Id for an object
     */
    public static String resolveEntityDefinitionId(String baseUrl, String sessionId, String objectApiName) {
        try {
            String q = 'SELECT Id FROM EntityDefinition WHERE QualifiedApiName = \'' + objectApiName + '\'';
            String endpoint = baseUrl + '/services/data/' + API_VERSION + '/tooling/query/?q=' + 
                             EncodingUtil.urlEncode(q, 'UTF-8');
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setHeader('Authorization', 'Bearer ' + sessionId);
            req.setHeader('Content-Type', 'application/json');
            req.setTimeout(60000);
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            if (res.getStatusCode() == 200) {
                Map<String,Object> pm = (Map<String,Object>) JSON.deserializeUntyped(res.getBody());
                List<Object> recs = (List<Object>) pm.get('records');
                if (recs != null && !recs.isEmpty()) {
                    Map<String,Object> r0 = (Map<String,Object>) recs[0];
                    return (String) r0.get('Id');
                }
            }
        } catch (Exception e) {
            System.debug('  ERROR resolveEntityDefinitionId: ' + e.getMessage());
        }
        return null;
    }
    
    /**
     * Try multiple approaches to get picklist values (with inactive values)
     */
    public static List<PicklistValueDetail> queryPicklistValuesWithFallback(
        String baseUrl,
        String sessionId,
        String objectApiName,
        String entityDefId,
        String fieldApiName
    ) {
        // 1) Try FieldDefinition (most reliable for getting inactive values)
        List<PicklistValueDetail> res = queryFieldDefinitionTooling(
            baseUrl, sessionId, objectApiName, fieldApiName
        );
        if (res != null && !res.isEmpty()) {
            return res;
        }
        
        // 2) Try CustomField with EntityDefinition.Id
        if (entityDefId != null) {
            res = queryCustomFieldTooling(baseUrl, sessionId, entityDefId, fieldApiName);
            if (res != null && !res.isEmpty()) {
                return res;
            }
        }
        
        // 3) Try CustomField with TableEnumOrId = objectApiName
        res = queryCustomFieldToolingTableEnum(baseUrl, sessionId, objectApiName, fieldApiName);
        if (res != null && !res.isEmpty()) {
            return res;
        }
        
        // 4) Final fallback: REST describe (may only return active values)
        res = queryRestDescribeForPicklist(baseUrl, sessionId, objectApiName, fieldApiName);
        if (res != null && !res.isEmpty()) {
            return res;
        }
        
        return new List<PicklistValueDetail>();
    }
    
    // --------------------------------------
    // Private HTTP callout methods
    // --------------------------------------
    
    private static List<PicklistValueDetail> queryFieldDefinitionTooling(
        String baseUrl, String sessionId, String objectApiName, String fieldApiName
    ) {
        List<PicklistValueDetail> results = new List<PicklistValueDetail>();
        
        try {
            String soql = 'SELECT Metadata FROM FieldDefinition WHERE EntityDefinition.QualifiedApiName = \'' + 
                         objectApiName + '\' AND QualifiedApiName = \'' + fieldApiName + '\'';
            String endpoint = baseUrl + '/services/data/' + API_VERSION + '/tooling/query/?q=' + 
                             EncodingUtil.urlEncode(soql, 'UTF-8');
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setHeader('Authorization', 'Bearer ' + sessionId);
            req.setHeader('Content-Type', 'application/json');
            req.setTimeout(60000);
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            if (res.getStatusCode() != 200) {
                return results;
            }
            
            Map<String,Object> body = (Map<String,Object>) JSON.deserializeUntyped(res.getBody());
            List<Object> recs = (List<Object>) body.get('records');
            if (recs == null || recs.isEmpty()) return results;
            
            Map<String,Object> rec = (Map<String,Object>) recs[0];
            Map<String,Object> metadata = (Map<String,Object>) rec.get('Metadata');
            if (metadata != null && metadata.containsKey('valueSet')) {
                parseValueSet(metadata, results);
            }
        } catch (Exception e) {
            System.debug('      ERROR queryFieldDefinitionTooling: ' + e.getMessage());
        }
        
        return results;
    }
    
    private static List<PicklistValueDetail> queryCustomFieldTooling(
        String baseUrl, String sessionId, String entityDefId, String fieldApiName
    ) {
        List<PicklistValueDetail> results = new List<PicklistValueDetail>();
        
        try {
            String devName = fieldApiName.endsWith('__c') ? 
                            fieldApiName.substring(0, fieldApiName.length()-3) : fieldApiName;
            
            String soql = 'SELECT Metadata FROM CustomField WHERE TableEnumOrId = \'' + 
                         entityDefId + '\' AND DeveloperName = \'' + devName + '\'';
            String endpoint = baseUrl + '/services/data/' + API_VERSION + '/tooling/query/?q=' + 
                             EncodingUtil.urlEncode(soql, 'UTF-8');
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setHeader('Authorization', 'Bearer ' + sessionId);
            req.setHeader('Content-Type', 'application/json');
            req.setTimeout(60000);
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            if (res.getStatusCode() != 200) {
                return results;
            }
            
            Map<String,Object> body = (Map<String,Object>) JSON.deserializeUntyped(res.getBody());
            List<Object> recs = (List<Object>) body.get('records');
            if (recs == null || recs.isEmpty()) return results;
            
            Map<String,Object> rec = (Map<String,Object>) recs[0];
            Map<String,Object> metadata = (Map<String,Object>) rec.get('Metadata');
            if (metadata != null && metadata.containsKey('valueSet')) {
                parseValueSet(metadata, results);
            }
        } catch (Exception e) {
            System.debug('      ERROR queryCustomFieldTooling: ' + e.getMessage());
        }
        
        return results;
    }
    
    private static List<PicklistValueDetail> queryCustomFieldToolingTableEnum(
        String baseUrl, String sessionId, String objectApiName, String fieldApiName
    ) {
        List<PicklistValueDetail> results = new List<PicklistValueDetail>();
        
        try {
            String devName = fieldApiName.endsWith('__c') ? 
                            fieldApiName.substring(0, fieldApiName.length()-3) : fieldApiName;
            
            String soql = 'SELECT Metadata FROM CustomField WHERE TableEnumOrId = \'' + 
                         objectApiName + '\' AND DeveloperName = \'' + devName + '\'';
            String endpoint = baseUrl + '/services/data/' + API_VERSION + '/tooling/query/?q=' + 
                             EncodingUtil.urlEncode(soql, 'UTF-8');
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setHeader('Authorization', 'Bearer ' + sessionId);
            req.setHeader('Content-Type', 'application/json');
            req.setTimeout(60000);
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            if (res.getStatusCode() != 200) {
                return results;
            }
            
            Map<String,Object> body = (Map<String,Object>) JSON.deserializeUntyped(res.getBody());
            List<Object> recs = (List<Object>) body.get('records');
            if (recs == null || recs.isEmpty()) return results;
            
            Map<String,Object> rec = (Map<String,Object>) recs[0];
            Map<String,Object> metadata = (Map<String,Object>) rec.get('Metadata');
            if (metadata != null && metadata.containsKey('valueSet')) {
                parseValueSet(metadata, results);
            }
        } catch (Exception e) {
            System.debug('      ERROR queryCustomFieldToolingTableEnum: ' + e.getMessage());
        }
        
        return results;
    }
    
    private static List<PicklistValueDetail> queryRestDescribeForPicklist(
        String baseUrl, String sessionId, String objectApiName, String fieldApiName
    ) {
        List<PicklistValueDetail> results = new List<PicklistValueDetail>();
        
        try {
            String endpoint = baseUrl + '/services/data/' + API_VERSION + '/sobjects/' + objectApiName + '/describe';
            
            HttpRequest req = new HttpRequest();
            req.setEndpoint(endpoint);
            req.setMethod('GET');
            req.setHeader('Authorization', 'Bearer ' + sessionId);
            req.setHeader('Content-Type', 'application/json');
            req.setTimeout(60000);
            
            Http http = new Http();
            HttpResponse res = http.send(req);
            
            if (res.getStatusCode() != 200) {
                return results;
            }
            
            Map<String,Object> body = (Map<String,Object>) JSON.deserializeUntyped(res.getBody());
            List<Object> fields = (List<Object>) body.get('fields');
            if (fields == null) return results;
            
            for (Object fObj : fields) {
                Map<String,Object> f = (Map<String,Object>) fObj;
                String name = (String) f.get('name');
                
                if (name != null && name.equalsIgnoreCase(fieldApiName)) {
                    List<Object> pickVals = (List<Object>) f.get('picklistValues');
                    if (pickVals != null) {
                        for (Object pvObj : pickVals) {
                            Map<String,Object> pv = (Map<String,Object>) pvObj;
                            PicklistValueDetail d = new PicklistValueDetail();
                            d.label = (String) pv.get('label');
                            d.value = (String) pv.get('value');
                            d.isActive = (Boolean) pv.get('active');
                            if (d.isActive == null) d.isActive = true;
                            results.add(d);
                        }
                    }
                    break;
                }
            }
        } catch (Exception e) {
            System.debug('      ERROR queryRestDescribeForPicklist: ' + e.getMessage());
        }
        
        return results;
    }
    
    private static void parseValueSet(Map<String,Object> metadata, List<PicklistValueDetail> out) {
        try {
            Object vs = metadata.get('valueSet');
            if (vs == null) return;
            
            Map<String,Object> valueSet = (Map<String,Object>) vs;
            
            // Check for valueSetDefinition structure
            if (valueSet.containsKey('valueSetDefinition')) {
                Map<String,Object> vsd = (Map<String,Object>) valueSet.get('valueSetDefinition');
                List<Object> values = (List<Object>) vsd.get('value');
                if (values != null) {
                    parseValueList(values, out);
                }
            } else if (valueSet.containsKey('value')) {
                // Alternate structure: direct value array
                List<Object> values = (List<Object>) valueSet.get('value');
                if (values != null) {
                    parseValueList(values, out);
                }
            }
        } catch (Exception e) {
            System.debug('      ERROR parseValueSet: ' + e.getMessage());
        }
    }
    
    private static void parseValueList(List<Object> values, List<PicklistValueDetail> out) {
        for (Object o : values) {
            Map<String,Object> m = (Map<String,Object>) o;
            PicklistValueDetail p = new PicklistValueDetail();
            p.label = (String) m.get('label');
            p.value = (String) m.get('valueName');
            if (p.value == null) p.value = (String) m.get('value');
            p.isActive = (Boolean) m.get('isActive');
            if (p.isActive == null) p.isActive = true;
            out.add(p);
        }
    }
}