import { ExportToCsv } from 'export-to-csv';
import { IJobApplication } from '../components/IJobApplication';
import { sp } from '@pnp/pnpjs';
import * as _ from 'lodash';
import * as moment from 'moment';

export interface ICSVFields {
  ApplicationId: number;
  ApplicationTitle: string;
  ApplicationCovernote: string;
  ApplicantName : string;
  ApplicantEmail : string;
  ApplicantCurrentRole : string;
  ApplicationCurrentManager : string;
  ApplicationDate: string;
  JobId: number;
  JobTitle: string;
  JobLevel: string;
  JobLocation: string;
  JobDescription: string;
  JobDeadline: string;
  JobAreaExpertise : string;
  JobTeam: string;
  JobArea: string;
  JobCreatedDate: string;
}
export default class CSVGenerator {

  private _buildCSVData = async(items : IJobApplication[]) => {
    //Get the related job details
    await items.forEach(async (item, i) =>{
      items[i].Job = await sp.web.lists.getByTitle('Jobs').items.getById(item.Job.Id).expand('Manager').select('Id', 'Title', 'Location', 'Deadline', 'Description', 'Created', 'Job_x0020_Level',
      'Manager/JobTitle', 'Manager/Name', 'Manager/EMail', 'Manager/Id','JobTags', 'Area', 'Team', 'Area_x0020_of_x0020_Expertise',
      'Manager/FirstName', 'Manager/LastName').get();
    });

    function sleep(ms) {
      return new Promise(resolve => setTimeout(resolve, ms));
    }
    await sleep(500);

    //flatten the object & use those fields
    const regex = /(?:\r\n|\r|\n)/g;
    let csvItems : ICSVFields[] =[]
    items.forEach(item =>{
      csvItems.push({
        ApplicationId: item.Id,
        ApplicationTitle: item.Title,
        ApplicationCovernote: item.Cover_x0020_Note? item.Cover_x0020_Note.replace(/<[^>]*>/g," ")
          .replace(regex, ' ').replace('&#160;' , ' '): '',
        ApplicantName : `${item.Author.FirstName} ${item.Author.LastName}`,
        ApplicantEmail : item.Author.EMail,
        ApplicantCurrentRole : item.Current_x0020_Role,
        ApplicationCurrentManager :  `${item.Current_x0020_Manager.FirstName} ${item.Current_x0020_Manager.LastName}`,
        ApplicationDate: moment(item.Created).format('YYYY-MM-DD').toString(),
        JobId: item.Job.Id,
        JobTitle: item.Job.Title,
        JobLevel: item.Job.Job_x0020_Level,
        JobLocation: item.Job.Location,
        JobDescription: item.Job.Description? item.Job.Description.replace(/<[^>]*>/g," ")
          .replace(regex, ' ').replace('&#160;' , ' '): '',
        JobDeadline: moment(item.Job.Deadline).format('YYYY-MM-DD').toString(),
        JobAreaExpertise : `${item.Job.Area_x0020_of_x0020_Expertise}`,
        JobTeam: `${item.Job.Team}`,
        JobArea: `${item.Job.Area}`,
        JobCreatedDate: moment(item.Job.Created).format('YYYY-MM-DD').toString()
      });
    });
    return csvItems;
  }

  public generateCSV = async(data) => {

    let csvItems = await this._buildCSVData(data);

    const options = {
      fieldSeparator: ',',
      quoteStrings: '"',
      decimalSeparator: '.',
      showLabels: true,
      showTitle: true,
      title: 'Job Applications',
      useTextFile: false,
      useBom: true,
      useKeysAsHeaders: true
    };


    const csvExporter = new ExportToCsv(options);

    csvExporter.generateCsv(csvItems);
  }

}

