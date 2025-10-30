/**
 * PicklistExportBatch - Batch job to export picklist metadata for multiple objects
 * Thin orchestration layer that delegates all work to PicklistExportHelper
 * 
 * SETUP:
 * 1) Deploy both classes: PicklistExportBatch and PicklistExportHelper
 * 2) Run from Developer Console:
 *      List<String> objects = new List<String>{'Account', 'Contact', 'Opportunity', 'Lead'};
 *      PicklistExportBatch batch = new PicklistExportBatch(objects);
 *      Database.executeBatch(batch, 1); // Process 1 object per batch execution
 * 
 * For 100+ objects, keep batch size small (1-5) depending on picklist count.
 */
public class PicklistExportBatch implements Database.Batchable<String>, Database.AllowsCallouts, Database.Stateful {
    
    // --------------------------------------
    // State variables
    // --------------------------------------
    private List<String> objectApiNames;
    private List<String> csvRows;
    private Integer totalValues = 0;
    private Integer totalInactive = 0;
    private String baseUrl;
    private String sessionId;
    
    // --------------------------------------
    // Constructor
    // --------------------------------------
    public PicklistExportBatch(List<String> objectNames) {
        this.objectApiNames = objectNames != null ? objectNames : new List<String>();
        this.csvRows = new List<String>();
        this.csvRows.add(PicklistExportHelper.getCsvHeaders());
        this.baseUrl = URL.getOrgDomainUrl().toExternalForm();
        this.sessionId = UserInfo.getSessionId();
        
        if (String.isBlank(this.sessionId)) {
            throw new IllegalArgumentException('SessionId unavailable. Run from Developer Console or authenticated context.');
        }
    }
    
    // --------------------------------------
    // Batchable interface methods
    // --------------------------------------
    public Iterable<String> start(Database.BatchableContext bc) {
        System.debug('=== PicklistExportBatch START ===');
        System.debug('Total objects to process: ' + objectApiNames.size());
        return objectApiNames;
    }
    
    public void execute(Database.BatchableContext bc, List<String> scope) {
        for (String objName : scope) {
            try {
                // Delegate all processing to helper
                PicklistExportHelper.ProcessingResult result = PicklistExportHelper.processObject(
                    objName, 
                    baseUrl, 
                    sessionId
                );
                
                // Accumulate results
                csvRows.addAll(result.csvRows);
                totalValues += result.valuesProcessed;
                totalInactive += result.inactiveValues;
                
            } catch (Exception e) {
                System.debug('ERROR processing object ' + objName + ': ' + e.getMessage() + '\n' + e.getStackTraceString());
            }

            // ✅ Heap safety: flush to file if usage exceeds 70%
            if (Limits.getHeapSize() > Limits.getLimitHeapSize() * 0.7) {
                System.debug('⚠️ Heap usage high (' + Limits.getHeapSize() + '/' + Limits.getLimitHeapSize() + '), flushing CSV to file...');
                PicklistExportHelper.createPartialCsvFile(csvRows);
                csvRows.clear();
                csvRows.add(PicklistExportHelper.getCsvHeaders());
            }
        }
    }
    
    public void finish(Database.BatchableContext bc) {
        System.debug('=== PicklistExportBatch FINISH ===');
        System.debug('Total values exported: ' + totalValues);
        System.debug('Total inactive values: ' + totalInactive);
        
        if (totalValues == 0) {
            System.debug('No picklist values found. No file created.');
            return;
        }
        
        // Delegate file creation to helper
        PicklistExportHelper.createFinalCsvFile(csvRows);
    }
}