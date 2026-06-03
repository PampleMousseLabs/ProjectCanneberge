<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>

      <tab id="tabPML"
           label="PampleMousse Labs">

        <group id="grpETL"
               label="Financial Data Engine">

          <!-- STAGED ETL REFRESH -->
          <button id="btnRefreshStagedModel"
                  label="Refresh Source Data"
                  size="large"
                  imageMso="RefreshAll"
                  onAction="RefreshSourceData_Click"/>

          <!-- BETA PIPELINE -->
          <button id="btnRefreshBeta"
                  label="β Refresh Beta"
                  size="normal"
                  imageMso="ChartLineMarkers"
                  onAction="RefreshBeta_Click"/>

          <!-- ETL LOG -->
          <button id="btnShowETLLog"
                  label="Show ETL Log"
                  size="normal"
                  imageMso="DatabaseToolsRelationships"
                  onAction="ShowETLLog_Click"/>

          <!-- CANCEL -->
          <button id="btnCancelRefresh"
                  label="Cancel ETL"
                  size="normal"
                  imageMso="CancelRequest"
                  onAction="CancelRefresh_Click"/>
             
                  

        </group>
        
        <group id="grpSlug" label="Slug Tools">

          <button id="btnSlugExtract"
                  label="Extract Market Slugs"
                  size="large"
                  imageMso="MicrosoftVisualFoxPro"
                  onAction="RunSlugButton"/>

        </group>

      </tab>

    </tabs>
  </ribbon>
</customUI>