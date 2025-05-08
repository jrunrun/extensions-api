import { DataSource } from '@tableau/extensions-api-types';

// Wrap everything in an anonymous function to avoid polluting the global namespace
(async () => {
  class DataSources {
    // Initialize any required properties in constructor
    constructor() {
      // No initialization needed at this time
    }

    /**
     * Refreshes the given dataSource
     * @param dataSource
     */
    private static async refreshDataSource(dataSource: DataSource) {
      await dataSource.refreshAsync();
      console.log(dataSource.name + ': Refreshed Successfully');
    }

    /**
     * Initializes the extension
     */
    public async initialize() {
      console.log('Waiting for DOM ready');
      // Wait for DOM content to be loaded
      if (document.readyState === 'loading') {
        await new Promise(resolve => {
          document.addEventListener('DOMContentLoaded', resolve);
        });
      }

      console.log('Initializing extension API');
      await tableau.extensions.initializeAsync();

      // Since dataSource info is attached to the worksheet, we will perform
      // one async call per worksheet to get every dataSource used in this
      // dashboard.  This demonstrates the use of Promise.all to combine
      // promises together and wait for each of them to resolve.
      const dataSourceFetchPromises: Array<Promise<DataSource[]>> = [];

      // To get dataSource info, first get the dashboard.
      const dashboard = tableau.extensions.dashboardContent.dashboard;
      // Then loop through each worksheet and get its dataSources, save promise for later.
      dashboard.worksheets.forEach(worksheet => dataSourceFetchPromises.push(worksheet.getDataSourcesAsync()));
      const fetchResults = await Promise.all(dataSourceFetchPromises);

      // Maps dataSource id to dataSource so we can keep track of unique dataSources.
      const dataSourcesCheck: Record<string, boolean> = {};
      const dashboardDataSources: DataSource[] = [];

      fetchResults.forEach(dss => {
        dss.forEach(ds => {
          if (!dataSourcesCheck[ds.id]) {
            // We've already seen it, skip it.
            dataSourcesCheck[ds.id] = true;
            dashboardDataSources.push(ds);
          }
        });
      });

      this.buildDataSourcesTable(dashboardDataSources);

      // This just modifies the UI by removing the loading banner and showing the dataSources table.
      const loadingElement = document.getElementById('loading');
      if (loadingElement) {
        loadingElement.classList.add('hidden');
      }
      const dataSourcesTable = document.getElementById('dataSourcesTable');
      if (dataSourcesTable) {
        dataSourcesTable.classList.remove('hidden');
        dataSourcesTable.classList.add('show');
      }
    }

    /**
     * Displays a modal dialog with more details about the given dataSource.
     * @param dataSource
     */
    private async showModal(dataSource: DataSource) {
      const modal = document.getElementById('infoModal');
      if (!modal) {
        return;
      }

      document.getElementById('nameDetail')!.textContent = dataSource.name;
      document.getElementById('idDetail')!.textContent = dataSource.id;
      document.getElementById('typeDetail')!.textContent = (dataSource.isExtract) ? 'Extract' : 'Live';

      // Loop through every field in the dataSource and concat it to a string.
      let fieldNamesStr = '';
      dataSource.fields.forEach(function(field) {
        fieldNamesStr += field.name + ', ';
      });
      // Slice off the last ", " for formatting.
      document.getElementById('fieldsDetail')!.textContent = fieldNamesStr.slice(0, -2);

      // Loop through each connection summary and list the connection's
      // name and type in the info field
      const connectionSummaries = await dataSource.getConnectionSummariesAsync();
      let connectionsStr = '';
      connectionSummaries.forEach(function(summary) {
        connectionsStr += summary.name + ': ' + summary.type + ', ';
      });
      // Slice of the last ", " for formatting.
      document.getElementById('connectionsDetail')!.textContent = connectionsStr.slice(0, -2);

      // Loop through each table that was used in creating this datasource
      const activeTables = await dataSource.getActiveTablesAsync();
      let tableStr = '';
      activeTables.forEach(function(table) {
        tableStr += table.name + ', ';
      });
      // Slice of the last ", " for formatting.
      document.getElementById('activeTablesDetail')!.textContent = tableStr.slice(0, -2);

      // Show the modal using Bootstrap 5's modal API
      // @ts-ignore
      const bsModal = new bootstrap.Modal(modal);
      bsModal.show();
    }

    /**
     * Constructs UI that displays all the dataSources in this dashboard
     * given a mapping from dataSourceId to dataSource objects.
     * @param dataSources
     */
    private buildDataSourcesTable(dataSources: DataSource[]) {
      const tableBody = document.querySelector('#dataSourcesTable > tbody');
      if (!tableBody) {
        return;
      }

      // Clear the table first.
      tableBody.innerHTML = '';

      // Add an entry to the dataSources table for each dataSource.
      for (const dataSource of dataSources) {
        const newRow = document.createElement('tr');
        const nameCell = document.createElement('td');
        const refreshCell = document.createElement('td');
        const infoCell = document.createElement('td');

        const refreshButton = document.createElement('button');
        refreshButton.innerHTML = 'Refresh Now';
        refreshButton.type = 'button';
        refreshButton.className = 'btn btn-primary';
        refreshButton.addEventListener('click', () => DataSources.refreshDataSource(dataSource));

        const infoSpan = document.createElement('span');
        infoSpan.className = 'bi bi-info-circle';  // Using Bootstrap Icons instead of Glyphicons
        infoSpan.style.cursor = 'pointer';  // Make it clear this is clickable
        infoSpan.addEventListener('click', () => this.showModal(dataSource));

        nameCell.textContent = dataSource.name;
        refreshCell.appendChild(refreshButton);
        infoCell.appendChild(infoSpan);

        newRow.appendChild(nameCell);
        newRow.appendChild(refreshCell);
        newRow.appendChild(infoCell);
        tableBody.appendChild(newRow);
      }
    }
  }

  console.log('Initializing DataSources extension.');
  await new DataSources().initialize();
})();
