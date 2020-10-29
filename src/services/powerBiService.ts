import * as $ from 'jquery';
import { isNullOrUndefined } from 'util';
import { Workspace, Dataset, Report, Dashboard } from '../models/models';

import AppSettings from './../appSettings';
import SpaAuthService from './SpaAuthService';

export default class PowerBiService {

  static apiRoot: string = "https://api.powerbi.com/v1.0/myorg/";

  static GetWorkspaces = async():Promise<Workspace[]> => {
    var accessToken: string = await SpaAuthService.getAccessToken();
    var restUrl = PowerBiService.apiRoot + "groups/";
    return $.ajax({
      url: restUrl,
      headers: {
        "Accept": "application/json;odata.metadata=minimal;",
        "Authorization": "Bearer " + accessToken
      }
    }).then(response => { return response.value });
  }

  static GetReports = async (workspaceId?: string): Promise<Report[]> => {
    console.log("GetReports: " + workspaceId)
    var accessToken: string = await SpaAuthService.getAccessToken();
    var restUrl = isNullOrUndefined(workspaceId) ? PowerBiService.apiRoot + "reports/" :
                                                   PowerBiService.apiRoot + "groups/" + workspaceId + "/reports/" ;
    return $.ajax({
      url: restUrl,
      headers: {
        "Accept": "application/json;odata.metadata=minimal;",
        "Authorization": "Bearer " + accessToken
      }
    }).then(response=>response.value);
  }

  static GetDashboards = async(workspaceId?: string):Promise<Dashboard[]> => {
    var accessToken: string = await SpaAuthService.getAccessToken();
    var restUrl = isNullOrUndefined(workspaceId) ? PowerBiService.apiRoot + "dashboards/" :
                                                   PowerBiService.apiRoot + "groups/" + workspaceId + "/dashboards/";
    return $.ajax({
      url: restUrl,
      headers: {
        "Accept": "application/json;odata.metadata=minimal;",
        "Authorization": "Bearer " + accessToken
      }
    }).then(response=>response.value);
  }

  static GetDatasets = async (workspaceId?: string): Promise<Dataset[]> => {
    var accessToken: string = await SpaAuthService.getAccessToken();
     var restUrl = isNullOrUndefined(workspaceId) ? PowerBiService.apiRoot + "datasets/" :
                                                    PowerBiService.apiRoot + "groups/" + workspaceId + "/datasets/" ;
   return $.ajax({
      url: restUrl,
      headers: {
        "Accept": "application/json;odata.metadata=minimal;",
        "Authorization": "Bearer " + accessToken
      }
    }).then(response=>response.value);
  }

}
