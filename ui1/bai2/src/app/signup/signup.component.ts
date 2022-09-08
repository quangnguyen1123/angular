import { Component, OnInit } from '@angular/core';
import { FormGroup ,FormBuilder} from '@angular/forms';
@Component({
  selector: 'app-signup',
  templateUrl: './signup.component.html',
  styleUrls: ['./signup.component.css']
})
export class SignupComponent implements OnInit {

  public SignupComponent !:FormGroup;

  public signupForm !:FormGroup;
  constructor(private formBuilder : FormBuilder){}
  ngOnInit(): void {
    this.signupForm = this.formBuilder.group({
      fullname:[""],
      email:[""],
      password:[""],
      mobile:[""],
    })
  }

}
