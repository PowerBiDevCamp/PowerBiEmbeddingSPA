import 'bootstrap/dist/css/bootstrap.min.css';
import 'bootstrap';
import "./app.css";

import * as $ from 'jquery';

import * as powerBiClient from "powerbi-client";
import * as pbimodels from "powerbi-models";

require('powerbi-models');
require('powerbi-client');

var powerbi: powerBiClient.service.Service = window.powerbi;

import SpaAuthService from './services/SpaAuthService';
import PowerBiService from './services/powerBiService'

import { Workspace, Dashboard, Report, Dataset } from './models/models';
import { ISettings } from 'embed';

$( () => {

  SpaAuthService.uiUpdateCallback = onAuthenticationCompleted;

  $("#login").on("click", async () => {
    await SpaAuthService.login();
  });

  $("#logout").on("click", () => {
    SpaAuthService.logout();
    refreshUi();
  });
  refreshUi();
});

var onAuthenticationCompleted = () => {
  refreshUi();
  initializeAppData();
}

var refreshUi = () => {
  if (SpaAuthService.userIsAuthenticated) {
    $("#user-greeting").text("Welcome " + SpaAuthService.userDisplayName);
    $("#login").hide()
    $("#logout").show();
    $("#view-anonymous").hide();
    $("#view-authenticated").show();
  }
  else {
   $("#user-greeting").text("");
    $("#login").show();
    $("#logout").hide();
    $("#view-anonymous").show();
    $("#view-authenticated").hide();
  }
}

var workspaces: Workspace[];

var initializeAppData = async () => {
  var workspaces: Workspace[] = await PowerBiService.GetWorkspaces();

  const urlParams = new URLSearchParams(window.location.search);
  const currentWorkspaceId = urlParams.get('workspaceId');
  console.log(currentWorkspaceId);
  var currentWorkspaceName: string;
  if (currentWorkspaceId) {
    console.log("Load App Workspace: " + currentWorkspaceId);
     currentWorkspaceName = workspaces.find(workspace => workspace.id == currentWorkspaceId).name
  }
  else {
    console.log("Load My Workspace");
    currentWorkspaceName = "My Workspace";
  }
  var workspaceSelector: JQuery = $("#workspace-selector");
  var workspacesList: JQuery = $("#workspaces-list");

  var linkMyWorkspace = $("<a>", { "href": "javascript:void(0);" })
    .text("My Workspace")
    .addClass("dropdown-item")
    .on("click", () => { 
      const params = new URLSearchParams(window.location.search);
      params.delete("workspaceId");
      window.history.replaceState({}, "", decodeURIComponent(`${window.location.pathname}?${params}`));
      workspaceSelector.text("My Workspace");
      loadWorkspace();
    });
  workspacesList.append(linkMyWorkspace);

  workspacesList.append($("<div>").addClass("dropdown-divider"))

  workspaces.forEach((workspace: Workspace) => {
    var workspaceId = workspace.id;
    var link = $("<a>", { "href": "javascript:void(0);" })
      .text(workspace.name)
      .addClass("dropdown-item")
      .click(() => { 
        const params = new URLSearchParams(window.location.search);
        params.set("workspaceId", workspace.id);
        window.history.replaceState({}, "", decodeURIComponent(`${window.location.pathname}?${params}`));
        var currentWorkspace: Workspace = workspaces.find(workspace => workspace.id == workspaceId);
        console.log("current workspace: " + currentWorkspace.name);
        workspaceSelector.text(currentWorkspace.name);
        loadWorkspace(currentWorkspace.id);
      });
    workspacesList.append(link);
  });
  workspaceSelector.text(currentWorkspaceName);
  loadWorkspace(currentWorkspaceId);
}

var loadWorkspace = async (workspaceId?: string) => {

   var embedContainer = document.getElementById('embed-container');
  // reset target div
  powerbi.reset(embedContainer);
  $("#embedding-instructions").show();

  console.log("loadWorkspace start: " + workspaceId);  

  var dashboards: Dashboard[] = await PowerBiService.GetDashboards(workspaceId);
  var reports: Report[] = await PowerBiService.GetReports(workspaceId);
  var datasets: Dataset[] = await PowerBiService.GetDatasets(workspaceId);

  var dashboardsList: JQuery = $("#dashboards-list").empty();
  if (dashboards.length == 0) {
    var listItem = $("<li>", { class: "nav-item empty-list" }).text("no dashboards");
    dashboardsList.append(listItem); 
  }
  else {
  dashboards.forEach((dashboard: Dashboard) => { 
    var link = $("<a>", { class: "nav-link" }).text(dashboard.displayName);
    link.attr("href", "JavaScript:void(0)")
    link.click(() => {
      embedDashboard(dashboard);
    });
    var listItem = $("<li>", { class: "nav-item" }).append(link);
    dashboardsList.append(listItem);
  });

  }

  var reportsList: JQuery = $("#reports-list").empty();
  if (reports.length == 0) {
    var listItem = $("<li>", { class: "nav-item empty-list" }).text("no reports");
    reportsList.append(listItem); 
  }
  else {
  reports.forEach((report: Report) => { 
     var link = $("<a>", { class: "nav-link" }).text(report.name);
    link.attr("href", "JavaScript:void(0)")
    link.click(() => {
      embedReport(report);
    });
    var listItem = $("<li>", { class: "nav-item" }).append(link);
    reportsList.append(listItem);
   });
  }
  
  var datasetsList: JQuery = $("#datasets-list").empty();
  if (datasets.length == 0) {
        var listItem = $("<li>", { class: "nav-item empty-list" }).text("no datasets");
    datasetsList.append(listItem); 
  }
  else {
  datasets.forEach((dataset: Dataset) => { 
     var link = $("<a>", { class: "nav-link" }).text(dataset.name);
    link.attr("href", "JavaScript:void(0)")
    link.click(() => {
      embedNewReport(dataset);
    });
    var listItem = $("<li>", { class: "nav-item" }).append(link);
    datasetsList.append(listItem);
  });

  }
  console.log("loadWorkspace end");  

}

var embedReport = async (report: Report, editMode: boolean = false) => {
  // Get a reference to the embedded report HTML element
  $("#embedding-instructions").hide();
  $("#toggle-edit").show();
  $("#embed-toolbar").show();
  $("#breadcrumb").text("Reports > " + report.name);

  var embedContainer = document.getElementById('embed-container');

  // reset target div
  powerbi.reset(embedContainer);

  // data required for embedding Power BI report
  var embedReportId = report.id;
  var embedUrl = report.embedUrl;
  var accessToken: string = await SpaAuthService.getAccessToken();

  // Get models object to access enums for embed configuration
  var models = pbimodels;

  var config: powerBiClient.IEmbedConfiguration = {
    type: 'report',
    id: embedReportId,
    embedUrl: embedUrl,
    accessToken: accessToken,
    tokenType: models.TokenType.Aad,
    permissions: models.Permissions.All,
    viewMode: (editMode ? models.ViewMode.Edit : models.ViewMode.View),
    settings: {
      visualRenderedEvents: true,
      useCustomSaveAsDialog: true,
      panes: {
        filters: { visible: true, expanded: false }
      }
    }
  };

  // Embed the report and display it within the div container
  var embeddedReport : powerBiClient.Report = powerbi.embed(embedContainer, config) as powerBiClient.Report ;

  console.log(embeddedReport);

  // toggle report between display mode and edit mode
    var viewMode = "view";
  $("#toggle-edit").off("click");  
  $("#toggle-edit").on("click", () => {
      viewMode = (viewMode == "view") ? "edit" : "view";
      embeddedReport.switchMode(viewMode);
      if (viewMode == "edit") {
        var settings : powerBiClient.IEmbedSettings = {
          panes: { filters: { visible: true} }
        };
        embeddedReport.updateSettings(settings);
      }
    });
  
    // command handler to enter full screen mode
  $("#full-screen").off("click");  
  $("#full-screen").on("click", () => {
      embeddedReport.fullscreen();
    });


  embeddedReport.off("saved");
  embeddedReport.on("saved", (args) => {
    console.log("saved", args);

  });

embeddedReport.off("saveAsTriggered")
  embeddedReport.on("saveAsTriggered", (args) => {
    console.log("saveAsTriggered", args);
  });
};

var embedDashboard = async (dashboard: Dashboard) => {

  $("#embedding-instructions").hide();
  $("#toggle-edit").hide();
  $("#embed-toolbar").show();
  $("#breadcrumb").text("Dashboards > " + dashboard.displayName);
  
  var embedContainer = document.getElementById('embed-container');

  // reset target div
  powerbi.reset(embedContainer);

  // data required for embedding Power BI dashboard
  var embedDashboardId = dashboard.id;
  var embedUrl = dashboard.embedUrl
  var accessToken = await SpaAuthService.getAccessToken();

  // Get models object to access enums for embed configuration
  var models = pbimodels;

  var config: any = {
    type: 'dashboard',
    id: embedDashboardId,
    embedUrl: embedUrl,
    accessToken: accessToken,
    tokenType: models.TokenType.Aad,
    pageView: "fitToWidth" // choices are "actualSize", "fitToWidth" or "oneColumn"
  };


  // Embed the report and display it within the div container.
  var embeddedDashboard = powerbi.embed(embedContainer, config);

    
  $("#full-screen").off("click");    
  $("#full-screen").on("click", () => {
    embeddedDashboard.fullscreen();
  });

}


var embedNewReport = async(dataset: Dataset) => {

  
  $("#embedding-instructions").hide();
  $("#toggle-edit").hide();
  $("#embed-toolbar").show();
  $("#breadcrumb").text("Datasets > " + dataset.name);
  

  var embedContainer = document.getElementById('embed-container');

  // reset target div
  powerbi.reset(embedContainer);

  // Get data required for embedding
  var embedDatasetId = dataset.id;
  var embedUrl = "https://app.powerbi.com/reportEmbed";
  var accessToken = await SpaAuthService.getAccessToken();

  // Get models object to access enums for embed configuration
  var models = pbimodels;

  var config: powerBiClient.IEmbedConfiguration = {
    datasetId: embedDatasetId,
    embedUrl: embedUrl,
    accessToken: accessToken,
    tokenType: models.TokenType.Aad,
    settings: { panes: {filters:{expanded:false}} }
  };

  console.log(config);

  // Embed the report and display it within the div container.
  var newReport = powerbi.createReport(embedContainer, config);

    
  $("#full-screen").off("click");    
  $("#full-screen").on("click", () => {
    newReport.fullscreen();
  });

  newReport.off("saved");
  newReport.on("saved", (event: any) => {

    var savedReport: Report = {
      id: event.detail.reportObjectId,
      name: event.detail.reportName,
      embedUrl: "https://app.powerbi.com/reportEmbed"
    };
     

    var link = $("<a>", { class: "nav-link" }).text(savedReport.name);
    link.attr("href", "JavaScript:void(0)")
    link.on("click", () => { embedReport(savedReport); });
    var listItem = $("<li>", { class: "nav-item" }).append(link);
    $("#reports-list").append(listItem);
  
    embedReport(savedReport, true);

  });
  

}
