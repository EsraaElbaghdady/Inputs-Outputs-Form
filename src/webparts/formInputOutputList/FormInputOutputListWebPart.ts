import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as $ from 'jquery';
import 'select2';
import 'select2/dist/css/select2.css';



export interface IIoFormWebPartProps {
  description: string;
}

export default class IoFormWebPart extends BaseClientSideWebPart<IIoFormWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `  
      <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">  
   <!-- Select2 CSS -->
<link href="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.13/css/select2.min.css" rel="stylesheet" />
<!-- Select2 JS -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.13/js/select2.min.js"></script>
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>


      <div class="form-container container mt-4">  
          <h2 class="text-center mb-4">Create New I/O </h2>  
          <div id="formMessage" class="alert d-none"></div>   
          <div><br></div>
          <form>
              <div class="form-row mb-3">  
                  <div class="col">  
                      <label for="Department">Department <span>*</span></label>  
                      <select id="Department" class="form-control rounded-input" placeholder="Insert Department">
                          <option></option>
                      </select>
                  </div>
                  <div class="col">  
          <label for="Sections">Sections </label>
              <select id="Sections" class="form-control  select2-multi" multiple="multiple" style="width: 100%;" placeholder="select Sections ">
                <option disabled>Select Sections</option>
              </select> 
      </div>     
<div class="form-row mb-3">  
  <div class="col">  
    <label for="Region">Region <span>*</span></label>  
    <select id="Region" class="form-control rounded-input">  
      <option value="EGY">EGY</option>  
      <option value="KSA">KSA</option>  
      <option value="Group">Group</option>  
    </select>  
  </div>
</div>
              </div> 
              <div class="form-row mb-4">
                  <div class="col">  
                      <label for="Level">Level<span>*</span></label>  
                      <select id="Level" class="form-control rounded-input"> 
                      <option ></option>  
                          <option value="Main Process">Main Process</option>  
                          <option value="Sub Process">Sub Process</option>  
                          <option value="Sub Sub Process">Sub Sub Process</option>  
                      </select>  
                  </div>
                    <div class="col">  
                <label for="MainProcess"> Process <span>*</span></label>
              <select id="MainProcess" class="form-control rounded-input" placeholder="Select Main Process">
               <option ></option>
              </select> 
          </div>  
              </div>
              <div class="form-row mb-3">  
                  <div class="col">  
                      <label for="Input">Input</label>  
                      <input type="text" id="Input" class="form-control rounded-input" /> 
                  </div>
                  <div class="col">  
                      <label for="Output">Output </label>  
                      <input type="text" id="Output" class="form-control rounded-input" /> 
                  </div>
              </div>
              <div class="form-row">  
                  <div class="col text-right">  
                      <button type="button" id="Cancel" class="btn btn-outline-secondary rounded-pill">Cancel</button>  
                      <button type="button" id="saveBtn" class="btn btn-brown rounded-pill">Save</button>  
                  </div>  
              </div>
          </form>  
      </div>
      <style>
         

        #formMessage {
          display: none; /* Initially hidden */
          margin-bottom: 20px;
          padding: 15px;
          border-radius: 5px;
          font-size: 16px;
        }

        .alert-success {
          background-color: #d4edda;
          color: #155724;
          border: 1px solid #c3e6cb;
        }

        .alert-warning {
          background-color: #fff3cd;
          color: #856404;
          border: 1px solid #ffeeba;
        }

        .d-none {
          display: none !important;
        }

        .btn-block {
          width: 100%;
        }

       .form-container {
            max-width: 900px;
            padding: 20px;
            margin: 0 auto; 
            flex-shrink: 0;
            border-radius: var(--Border-Raduis-Raduis-md, 10px);
            border-top: var(--Border-Raduis-Raduis-md, 10px) solid var(--gradients-Golden-Gradiant, #AE8C67);
            border-bottom: 1px solid #9C7D5C;
            border-right: 1px solid #9C7D5C;
            border-left: 1px solid #9C7D5C;
            background: #FFF;
            box-shadow: 0px 1px 2px 0px Global.Color.NeutralShadowAmbientBrand, 0px 0px 2px 0px Global.Color.NeutralShadowAmbient;
        } 


        .form-container h2 {  
            color: #6e5033;  
            text-align: center;  
            margin-bottom: 20px;   
        }  
        .form-container span{
        color: #FF6363;
        }

        .form-control, .rounded-input {  
            border-radius: 12px;  
        }  

        label {  
            font-weight: bold;  
            color: #6e5033;  
        }  

        small {  
            color: #999;  
        }  

        .btn-outline-secondary {  
            border: 1px solid #6A6265;  
            color: #6A6265;  
            background-color: transparent;  
        }  

        .btn-brown {  
            background-color: #9C7D5C;  
            color: white;  
        }  

        .btn {  
            padding: 10px 20px;  
            margin-left: 10px;  
        }  

        .rounded-pill {  
            border-radius: 15px;  
        }

        button-saveBtn:hover {
          background-color: #8b5a30;
        }

        .btn-cancel:hover {
          background-color: #bbb;
        } 

    
      </style>
    `;

    this.setButtonEventHandler();
    this.populateLookupFields();
  }

  private setButtonEventHandler(): void {
    const saveButton: HTMLElement | null = this.domElement.querySelector('#saveBtn');
    const cancelButton: HTMLElement | null = this.domElement.querySelector('#Cancel');
    
    if (saveButton) {
      saveButton.addEventListener('click', () => this.saveForm());
    }
    
    if (cancelButton) {
      cancelButton.addEventListener('click', () => { 
        window.location.href = 'https://andalusiagroupegypt.sharepoint.com/sites/Apps/Process/SitePages/Process-Page.aspx';
      });
    }
  }

  private populateLookupFields(): void {
    const departmentSelect = this.domElement.querySelector('#Department') as HTMLSelectElement;
   
    const levelSelect = this.domElement.querySelector('#Level') as HTMLSelectElement;
    

    // Fetch Departments and populate dropdown
    this.getLookupItems('Department').then(departments => {
        departmentSelect.innerHTML = '<option>Select Department</option>';
        departments.forEach(department => {
            departmentSelect.innerHTML += `<option value="${department.Id}">${department.Title}</option>`;
        });

        // Attach department change event
        departmentSelect.addEventListener('change', this.onDepartmentChange.bind(this));
    });

    // Attach level change event to trigger process population
    levelSelect.addEventListener('change', () => {
        const selectedDepartment = departmentSelect.value;
        const selectedLevel = levelSelect.value;


        if (selectedDepartment && selectedLevel) {
          console.log(selectedLevel);
          
            this.loadProcesses(selectedDepartment, selectedLevel);
        }
    });

    
}

private onDepartmentChange(): void {
  const departmentSelect = this.domElement.querySelector('#Department') as HTMLSelectElement;
  const sectionsSelect = this.domElement.querySelector('#Sections') as HTMLSelectElement;
  
  sectionsSelect.innerHTML = '<option>Select Sections</option>'; // Clear previous options
  this.getLookupItemsfiltered('Sections', departmentSelect.value).then(sections => {
      sections.forEach(section => {
          const option = document.createElement('option');
          option.value = section.Id;
          option.textContent = section.Title;
          sectionsSelect.appendChild(option);
      });
  });

  setTimeout(() => {
    ($('.select2-multi') as any).select2({
      placeholder: '',
      allowClear: true,
      closeOnSelect: false, 
      templateSelection: (selectedOption: any) => {
        return `<span style="color: white; background-color:#9C7D5C; padding: 3px 5px; border-radius: 5px;">${selectedOption.text}</span>`;
      },
      templateResult: (option: any) => {
        return `<span>${option.text}</span>`;
      },
      escapeMarkup: (markup: string) => markup, 
    });
  }, 0);
}


private loadProcesses(departmentId: string, level: string): void {
    const processSelect = this.domElement.querySelector('#MainProcess') as HTMLSelectElement;
console.log(level);

    // Clear previous options
    processSelect.innerHTML = '<option>Select Main Process</option>';
console.log(level);

    // Fetch filtered processes based on department and level
    this.getLookupItemsfiltered('Process', departmentId, level).then(processes => {
        processes.forEach(process => {
            processSelect.innerHTML += `<option value="${process.Id}">${process.Title}</option>`;
        });
    });
}

  private getLookupItems(listTitle: string): Promise<any[]> {
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$top=5000`;
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => response.json())
      .then(data => data.value);
  }

  private getLookupItemsfiltered(listTitle: string, department: string,level?: string): Promise<any[]> {
    const dept = Number(department);
    let url = '';
    console.log(level);

    if (listTitle === 'Sections') {
        url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$top=5000&$filter=DepartmentName/Id eq ${dept}`;
    } else if (listTitle === 'Process') {
console.log(level);

        url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$top=5000&$filter=Department/Id eq ${dept} and ProcessLevel eq '${level}'`;

    } else {
        return Promise.reject(new Error("Invalid list title"));
    }

    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => response.json())
        .then(data => {
            console.log('Data from SharePoint API:', data);
            return data.value;
        });
}


  

private saveForm(): void {
  const department = (this.domElement.querySelector('#Department') as HTMLSelectElement).value;
  const section = (this.domElement.querySelector('#Sections') as HTMLSelectElement).value;
  const level=(this.domElement.querySelector('#Level') as HTMLSelectElement).value;
  const process = (this.domElement.querySelector('#MainProcess') as HTMLSelectElement).value;
  const input = (this.domElement.querySelector('#Input') as HTMLInputElement).value;
  const output = (this.domElement.querySelector('#Output') as HTMLInputElement).value;
  const region = (this.domElement.querySelector('#Region') as HTMLSelectElement).value;
  if (!department || !section || !process  || !output || !region || !level) {
    this.showMessage('Please complete all fields', 'error');
    return;
  }
  const formData = {
    Department: department,
    Section: section,
    Process: process,
    Input: input,
    Output: output,
    Region: region,
  };

  console.log('Form data to save:', formData);

  // Save data to a SharePoint list
  this.saveDataToSharePoint(formData).then(() => {
    console.log('Form data saved successfully');
    this.showMessage('Form data saved successfully', 'success');
  }).catch(error => {
    console.error('Error saving form data:', error);
    this.showMessage('Error saving form data', 'error');
  });
}

private saveDataToSharePoint(formData: any): Promise<void> {
  const listTitle = 'Inputs/Outputs'; // Replace with the name of your SharePoint list
  const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items`;

  const body={
    
    DepartmentId: parseInt( formData.Department),
    SectionNameId:parseInt (formData.Section),
    ProcessNameId: parseInt(formData.Process),
    ProcessInput: formData.Input,
    ProcessOutput: formData.Output,
    Region: formData.Region,
  };

  return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=nometadata',
      'odata-version': ''
    },
    body: JSON.stringify(body)
  }).then((response: SPHttpClientResponse) => {
    if (!response.ok) {
      throw new Error('Failed to save data to SharePoint');
    }
    return response.json();
  }).then(() => {
    this.showMessage('Form data saved successfully', 'success');
    this.resetForm(); // Call the resetForm function here
  });
}

private showMessage(message: string, type: 'success' | 'error'): void {
  const messageDiv = this.domElement.querySelector('#formMessage') as HTMLDivElement;
  messageDiv.textContent = message;
  messageDiv.classList.remove('d-none', 'alert-success', 'alert-warning');
  messageDiv.classList.add(type === 'success' ? 'alert-success' : 'alert-warning');
  messageDiv.style.display = 'block';
}
private resetForm(): void {
  const departmentSelect = this.domElement.querySelector('#Department') as HTMLSelectElement;
  const sectionsSelect = this.domElement.querySelector('#Sections') as HTMLSelectElement;
  const levelSelect = this.domElement.querySelector('#Level') as HTMLSelectElement;
  const processSelect = this.domElement.querySelector('#MainProcess') as HTMLSelectElement;
  const inputField = this.domElement.querySelector('#Input') as HTMLInputElement;
  const outputField = this.domElement.querySelector('#Output') as HTMLInputElement;
  const regionSelect = this.domElement.querySelector('#Region') as HTMLSelectElement;

  departmentSelect.value = '';
  sectionsSelect.value = '';
  levelSelect.value = '';
  processSelect.value = '';
  inputField.value = '';
  outputField.value = '';
  regionSelect.value = '';

  // Clear previous options for sections and process
  sectionsSelect.innerHTML = '<option>Select Sections</option>';
  processSelect.innerHTML = '<option>Select Main Process</option>';
}
}