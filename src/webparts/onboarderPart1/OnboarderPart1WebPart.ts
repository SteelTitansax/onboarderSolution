import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import styles from './OnboarderPart1WebPart.module.scss';
import * as strings from 'OnboarderPart1WebPartStrings';

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

          <input type="text" id="Name" placeholder="Name" name="Name" />
          <input type="text" id="Surname" placeholder="Surname" name="Surname" />
          <input type="text" id="Role" placeholder="Role" name="Role" />
          <input type="text" id="OnboardingDate" placeholder="OnboardingDate" name="OnboardingDate" />
          <input type="text" id="MaterialDelivered" placeholder="MaterialDelivered" name="MaterialDelivered" />
          <input type="text" id="OfficeTraining" placeholder="OfficeTraining" name="OfficeTraining" />
          <input type="text" id="RoleTraning" placeholder="RoleTraning" name="RoleTraning" />
          <input type="text" id="SecurityCardDelivered" placeholder="SecurityCardDelivered" name="SecurityCardDelivered" />
          <input type="text" id="OnboardingDocumentsSigned" placeholder="OnboardingDocumentsSigned" name="OnboardingDocumentsSigned" />
        
        </div>
        
        
        <!-- Approver -->
        <div class=${styles.formBody} id="Approver">
          <p> Please the fill the Approver Person </p>
          <input type="text" id="approver" placeholder="Contact Person" name="Contact Person" />

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

  }
    private _initialLayout(): void {
      // Initial Layout 
      var step1 =  document.getElementById('BasicDetails');
      if(step1)
      {
        step1.style.display = 'block';  
      }

      var step2 =  document.getElementById('Approver');
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
  
          var step2 =  document.getElementById("Approver");
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

        var step2 =  document.getElementById("Approver");
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
  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      
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
