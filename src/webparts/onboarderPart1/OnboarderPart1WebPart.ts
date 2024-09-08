import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import styles from './OnboarderPart1WebPart.module.scss';
import * as strings from 'OnboarderPart1WebPartStrings';
import {ISPHttpClientOptions, SPHttpClient,SPHttpClientResponse} from '@microsoft/sp-http';

export interface IOnboarderPart1WebPartProps {
  description: string;
}

export default class OnboarderPart1WebPart extends BaseClientSideWebPart<IOnboarderPart1WebPartProps> {

 

  public render(): void {
    this.domElement.innerHTML = `
    <div>

        
    
        <!-- Basic details -->
        <div class=${styles.formBody} id="BasicDetails">

          <p> Please the fill in the form basic details </p>

          <input type="text" id="name" placeholder="Name" name="Name" />
          <input type="text" id="surname" placeholder="Surname" name="Surname" />
          <input type="text" id="role" placeholder="Role" name="Role" />
          <input type="text" id="onboardingDate" placeholder="OnboardingDate" name="OnboardingDate" />
          <input type="text" id="materialDelivered" placeholder="MaterialDelivered" name="MaterialDelivered" />
          <input type="text" id="officeTraining" placeholder="OfficeTraining" name="OfficeTraining" />
          <input type="text" id="roleTraning" placeholder="RoleTraning" name="RoleTraning" />
          <input type="text" id="securityCardDelivered" placeholder="SecurityCardDelivered" name="SecurityCardDelivered" />
          <input type="text" id="onboardingDocumentsSigned" placeholder="OnboardingDocumentsSigned" name="OnboardingDocumentsSigned" />
        
        </div>
        
        
        <!-- Approver -->
        <div class=${styles.formBody} id="approver">
          <p> Please the fill the Approver Person </p>
          <input type="text" id="approverName" placeholder="Contact Person" name="Contact Person" />

        </div>
        
        
    
      <br/>

        <div>
            <!-- Submit Button -->
           
                  
            <!-- PaginationButtons -->
            <div class=${styles.buttonsLeft} id ="NavigationButtonsLeft">
            <input  type="button" id="BackBasicDetails" value="Back - Basic Details"></input>  
            

            </div>
            <div class=${styles.buttonsRight} id ="NavigationButtonsRight">
            <input type="button"  id="approverBtn" value="Next - Approvers"></input>
            <input type="button" id="BttnEmp" value="Submit"></input> 
            </div>
        </div>
      
      
      </div>`;
    this._initialLayout();
    this._showNextStep();
    this._BackBasicDetails();
    this._bindSave();


  }
    private _initialLayout(): void {
      // Initial Layout 
      var step1 =  document.getElementById('BasicDetails');
      if(step1)
      {
        step1.style.display = 'block';  
      }

      var step2 =  document.getElementById('approver');
      if(step2)
      {
        step2.style.display = 'none';  
      }

      var BttnEmp =  document.getElementById('BttnEmp');
      if(BttnEmp)
      {
        BttnEmp.style.display = 'none';  
      }
      
      
    }

  private _showNextStep(): void  {
      
      const NextApprovers = this.domElement.querySelector('#approverBtn');
      if (NextApprovers) {

        NextApprovers.addEventListener('click', () => { 
          var step1 =  document.getElementById("BasicDetails");
          if(step1)
          {
            step1.style.display = 'none';  
          }
  
          var step2 =  document.getElementById("approver");
          if(step2)
          {
            step2.style.display = 'block';  
          }
          var approverBtn =  document.getElementById("approverBtn");
          if(approverBtn)
          {
            approverBtn.style.display = 'none';  
          }
          
          var BttnEmp =  document.getElementById("BttnEmp");
          if(BttnEmp)
          {
            BttnEmp.style.display = 'block';  
          }
          
      });
          
    } else {
      console.error("Button element '#NextApprovers' not found.");
    }
    
  }
  
  private _BackBasicDetails(): void  {
      
    const BackBasicDetails = this.domElement.querySelector('#BackBasicDetails');
    if (BackBasicDetails) {

      BackBasicDetails.addEventListener('click', () => { 
        var step1 =  document.getElementById("BasicDetails");
        if(step1)
        {
          step1.style.display = 'block';  
        }

        var step2 =  document.getElementById("approver");
        if(step2)
        {
          step2.style.display = 'none';  
        }
        var approverBtn =  document.getElementById("approverBtn");
        if(approverBtn)
        {
          approverBtn.style.display = 'block';  
        }
        
        var BttnEmp =  document.getElementById("BttnEmp");
        if(BttnEmp)
        {
          BttnEmp.style.display = 'none';  
        }
        
    });
        
  } else {
    console.error("Button element '#BackBasicDetails' not found.");
  }
}
  
  
  private _bindSave(): void {

    const button = this.domElement.querySelector('#BttnEmp');
    if (button) {
        button.addEventListener('click', () => { this.addListItem(); });
    } else {
        console.error("Button element '#BttnEmp' not found.");
    }
  
  }

    private addListItem(): void {
      var name = (document.getElementById("name") as HTMLInputElement).value;
      var surname = (document.getElementById("surname") as HTMLInputElement).value;
      var role = (document.getElementById("role") as HTMLInputElement).value;
      var onboardingDate = (document.getElementById("onboardingDate") as HTMLInputElement).value;
      var materialDelivered = (document.getElementById("materialDelivered") as HTMLInputElement).value;
      var officeTraining = (document.getElementById("officeTraining") as HTMLInputElement).value;
      var roleTraning = (document.getElementById("roleTraning") as HTMLInputElement).value;
      var securityCardDelivered = (document.getElementById("securityCardDelivered") as HTMLInputElement).value;
      var onboardingDocumentsSigned = (document.getElementById("onboardingDocumentsSigned") as HTMLInputElement).value;


      const siteUrl: string = "https://t8656.sharepoint.com/sites/Sharepoint_Interaction/_api/web/lists/getbytitle('PoC_SharepointInteraction')/items";
      const itemBody: any = {
          "Title": name+"_"+surname, 
          "Name": name,
          "Surame": surname,
          "Role": role,
          "OnboardingDate": onboardingDate,
          "MaterialDelivered": materialDelivered,
          "OfficeTraining_x003f_": officeTraining,
          "RoleTraning_x003f_": roleTraning,
          "SecurityCardDelivered_x003f_": securityCardDelivered,
          "OnboardingDocumentsSigned_x003f_": onboardingDocumentsSigned
      };
      const spHttpClientOptions: ISPHttpClientOptions = {
          "body": JSON.stringify(itemBody)
      };

      this.context.spHttpClient.post(siteUrl, SPHttpClient.configurations.v1, spHttpClientOptions)
          .then((response: SPHttpClientResponse) => {
              if (response.ok) {
                  // Llamar a la función para añadir ítems en la lista auxiliar
                  this.addAuxListItems(name+"_"+surname);
                  window.location.href = "https://t8656.sharepoint.com/sites/Sharepoint_Interaction/SitePages/Onboarder_Tracker.aspx?workerName=" + name+"_"+surname;
              } else {
                  console.error("Error adding list item:", response.statusText);
                  alert("Error adding item.");
              }
          })
          .catch((error) => {
              console.error("Error adding list item:", error);
              alert("Error adding item.");
          });
  }
  
  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      
    });
  }

    private addAuxListItems(fullname : string): void {
      const siteUrl: string = "https://t8656.sharepoint.com/sites/Sharepoint_Interaction/_api/web/lists/getbytitle('PoC_SharepointInteractionAux')/items";
      
      // Datos para los tres ítems
      const items = [
          { "Title": fullname, "Name": (document.getElementById("approverName") as HTMLInputElement).value, "Role": "Manager","Status": "Pending"},
          { "Title": fullname, "Name": "Manuel Portero", "Role": "HelpDesk","Status": "Pending"},
          { "Title": fullname, "Name": "Rosa Hernandez", "Role": "FrontDesk","Status": "Pending"}
      ];

      items.forEach(itemBody => {
          const spHttpClientOptions: ISPHttpClientOptions = {
              "body": JSON.stringify(itemBody)
          };

          this.context.spHttpClient.post(siteUrl, SPHttpClient.configurations.v1, spHttpClientOptions)
              .then((response: SPHttpClientResponse) => {
                      console.log("adding aux list item:", response.statusText);
              })
              .catch((error) => {
                  console.error("Error adding aux list item:", error);
              });
      });
  }


  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
