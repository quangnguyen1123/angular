import { Component, OnInit } from '@angular/core';
import { SharedService } from 'src/app/shared.service';
import * as XLSX from 'xlsx';
@Component({
  selector: 'app-show-dep',
  templateUrl: './show-dep.component.html',
  styleUrls: ['./show-dep.component.css']
})
export class ShowDepComponent implements OnInit {
  title = 'angular-app';
  fileName= 'ExcelSheet.xlsx';

  constructor(private service:SharedService) { }

  DepartmentList:any=[];

  ModalTitle:string |undefined;
  ActivateAddEditDepComp:boolean=false;
  dep:any;

  ngOnInit(): void {
    this.refreshDepList();
  }

  addClick(){
    this.dep={
      DepartmentId:0,
      DepartmentName:""
    }
    this.ModalTitle= "Add Department";
    this.ActivateAddEditDepComp=true;
  }

  editClick(item: any){
    this.dep= item;
    this.ModalTitle="Edit Dapertment";
    this.ActivateAddEditDepComp=true;
  }
  closeClick(){
    this.ActivateAddEditDepComp=false;
    this.refreshDepList();
  }
  deleteClick(item:any){
    if(confirm('are you sure??')){
      this.service.deleteDepartment(item.DepartmentId).subscribe(data=>{
        alert(data.toString());
        this.refreshDepList();
      })
    }
  }

  refreshDepList(){
    this.service.getDepList().subscribe(data=>{
      this.DepartmentList = data;
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
