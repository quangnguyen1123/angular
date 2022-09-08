import { Component, OnInit } from '@angular/core';
import { SharedService } from 'src/app/shared.service';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-show-emp',
  templateUrl: './show-emp.component.html',
  styleUrls: ['./show-emp.component.css']
})
export class ShowEmpComponent implements OnInit {

  title = 'angular-app';
  fileName= 'ExcelSheet.xlsx';
  
  constructor(private service:SharedService) { }

  EmployeeList:any=[];

  ModalTitle:string |undefined;
  ActivateAddEditEmpComp:boolean=false;
  emp:any;

  ngOnInit(): void {
    this.refreshEmpList();
  }

  addClick(){
    this.emp={
      EmployeeId: 0,
      EmployeeName:"",
      Department:"",
      DateOfJoining:"",
      PhotoFileName:"anonymous.png"
    }
    this.ModalTitle= "Add Employee";
    this.ActivateAddEditEmpComp=true;
  }

  editClick(item: any){
    this.emp= item;
    this.ModalTitle="Edit Employee";
    this.ActivateAddEditEmpComp=true;
  }
  closeClick(){
    this.ActivateAddEditEmpComp=false;
    this.refreshEmpList();
  }
  deleteClick(item:any){
    if(confirm('are you sure??')){
      this.service.deleteDepartment(item.EmployeeId).subscribe(data=>{
        alert(data.toString());
        this.refreshEmpList();
      })
    }
  }

  refreshEmpList(){
    this.service.getEmpList().subscribe(data=>{
      this.EmployeeList = data;
    });
  }


  exportexcel(): void
  {
    /* pass here the table id */
    let element = document.getElementById('quang');
    const ws: XLSX.WorkSheet =XLSX.utils.table_to_sheet(element);
 
    /* generate workbook and add the worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
 
    /* save to file */  
    XLSX.writeFile(wb, this.fileName);
 
  }

}
