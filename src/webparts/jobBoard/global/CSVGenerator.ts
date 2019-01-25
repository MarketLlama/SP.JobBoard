import { ExportToCsv } from 'export-to-csv';
export default class CVSGenerator {

  public generateCSV = (data) => {

    const options = {
      fieldSeparator: ',',
      quoteStrings: '"',
      decimalSeparator: '.',
      showLabels: true,
      showTitle: true,
      title: 'Job Applications',
      useTextFile: false,
      useBom: true,
      useKeysAsHeaders: true,
      // headers: ['Column 1', 'Column 2', etc...] <-- Won't work with useKeysAsHeaders present!
    };


    const csvExporter = new ExportToCsv(options);

    csvExporter.generateCsv(data);
  }

}

