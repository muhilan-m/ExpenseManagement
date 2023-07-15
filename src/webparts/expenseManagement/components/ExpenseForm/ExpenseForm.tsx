import { Dropdown, IDropdownOption, PrimaryButton, TextField } from '@fluentui/react';
import * as React from 'react';
//import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IExpenseFormProps } from './IExpenseFormProps';
//import { BaseComponentContext } from '@microsoft/sp-component-base';
// import { Logger, LogLevel } from "@pnp/logging";
import { spfi, SPFI } from "@pnp/sp";
import { getSP } from "../../pnpjsConfig";
import { Caching } from "@pnp/queryable";
//import { ToastContainer, toast } from 'react-toastify';
 // import 'react-toastify/dist/ReactToastify.css';
// import { IItemUpdateResult } from "@pnp/sp/items";
interface IExpenseFormData {
  user: { name: string, email: string };
  department: IDropdownOption;
  projectName: string;
  Expense: string;
  Remarks: string;
}

const ExpenseForm = (props: IExpenseFormProps) => {
  //const baseComponentContext: BaseComponentContext = props.spContext;

  const _sp: SPFI = getSP();

  const [formData, setFormData] = React.useState<IExpenseFormData>({
    user: null,
    department: null,
    projectName: '',
    Expense: '',
    Remarks: ''
  });

  const [departments, setDepartments] = React.useState<IDropdownOption[]>([]);

  React.useEffect(() => {
    const spCache = spfi(_sp).using(Caching({ store: "session" }));
    spCache.web.lists.getByTitle("Departments").items.select("Title,Id")().then((val) => {
      setDepartments(val.map((item)=>({'key':item.Id,'text':item.Title})))
    });

    setFormData({ ...formData, user: { name: props.spContext.pageContext.user.displayName, email: props.spContext.pageContext.user.email } })

    
  }, [])

  const handleChange = (event: any, fieldName: string) => {
    setFormData({ ...formData, [fieldName]: event.target.value });
  };

  const handleDepartmentChange=(event:any)=>{
    setFormData({ ...formData, department: event});
  }

  const handleSubmit = async (event: any) => {
    event.preventDefault();
    // Handle form submission
    const user = await _sp.web.ensureUser(formData.user.email);
    const userId = user.data.Id;

    _sp.web.lists.getByTitle("ExpenseDetails").items.add({
      NameId:userId,
      ProjectName:formData.projectName,
      Expense:formData.Expense,
      Remarks:formData.Remarks,
      DepartmentId:formData.department.key
    }).then(()=>{
      
      alert("Item Created");
      setFormData({...formData,
        department: null,
        projectName: '',
        Expense: '',
        Remarks: ''
      })
    })

    console.log(formData);
  };

  return (
    <>
      <div>ExpenseForm</div>
      <form onSubmit={handleSubmit}>
        <TextField
          label="Name"
          value={formData?.user?.name}
          onChange={(e) => handleChange(e, 'name')}
          required
          disabled
        />
        <Dropdown
          label="Department"
          options={departments}
          selectedKey={formData.department ? formData.department.key : undefined}
          onChanged={(e) => handleDepartmentChange(e)}
          required
        />
        <TextField
          label="Project Name"
          value={formData.projectName}
          onChange={(e) => handleChange(e, 'projectName')}
          required
        />
        <TextField
          label="Expense"
          value={formData.Expense}
          onChange={(e) => handleChange(e, 'Expense')}
          type='number'
          required
        />
        <TextField
          label="Remarks"
          value={formData.Remarks}
          onChange={(e) => handleChange(e, 'Remarks')}
          multiline
        />
        <PrimaryButton type="submit">Submit</PrimaryButton>
      </form>
      {/* <ToastContainer /> */}
    </>
  )
}

export default ExpenseForm