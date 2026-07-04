<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>

      <tab id="tabPML"
           label="PampleMousse Labs">

        <group id="grpETL"
               label="Financial Data Engine">

          <button id="btnRefreshStagedModel"
                  label="Refresh Source Data"
                  size="large"
                  imageMso="RefreshAll"
                  onAction="RefreshSourceData_Click"/>

          <button id="btnRefreshBeta"
                  label="β Refresh Beta"
                  size="normal"
                  imageMso="ChartLineMarkers"
                  onAction="RefreshBeta_Click"/>

          <button id="btnRefreshForwardEst"
                  label="Refresh Forward Est"
                  size="normal"
                  imageMso="TimeScaleMenu"
                  onAction="RefreshForwardEst_Click"/>

          <button id="btnCancelRefresh"
                  label="Cancel ETL"
                  size="normal"
                  imageMso="CancelRequest"
                  onAction="CancelRefresh_Click"/>

        </group>

        <group id="grpDiagnostics"
               label="Diagnostics Tools">

          <button id="btnShowETLLog"
                  label="Show ETL Log"
                  size="large"
                  imageMso="DatabaseRelationships"
                  onAction="ShowETLLog_Click"/>

          <button id="btnClearETLLog"
                  label="Clear ETL Log"
                  size="normal"
                  imageMso="InkEraser"
                  onAction="ClearETLLog_Click"/>

          <button id="btnTestConnection"
                  label="Test Connection"
                  size="normal"
                  imageMso="GetExternalDataFromWeb"
                  onAction="TestConnection_Click"/>

          <button id="btnRunSummary"
                  label="Run Summary"
                  size="normal"
                  imageMso="PropertySheet"
                  onAction="ShowRunSummary_Click"/>                 
       
        </group>         
        
        <group id="grpSheetTools"
               label="Sheet Tools">
                  
          <button id="btnRefreshCompChart"
                  label="Refresh Comp Chart"
                  size="large"
                  imageMso="ChartErrorBars"
                  onAction="RefreshCompChart_Click"/>

        </group>
            
        <group id="grpDevTools" label="Dev Tools">
            
            <button id="btnExportCode"
                  label="Export Code"
                  imageMso="ExportTextFile"
                  size="large"
                  onAction="ExportCode_Click"
                  supertip="Export all VBA modules and Power Query M code to the local repo folder (\vba\ and \power-query\)."/>
        </group>                     

      </tab>

    </tabs>
  </ribbon>
</customUI>