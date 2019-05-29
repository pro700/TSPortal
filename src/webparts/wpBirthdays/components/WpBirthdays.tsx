import "core-js/modules/es6.promise"; 
import "core-js/modules/es6.array.iterator.js"; 
import "core-js/modules/es6.array.from.js"; 
import "whatwg-fetch";
import "es6-map/implement";

import * as React from 'react';
import * as Reactdom from 'react-dom';
import styles from './WpBirthdays.module.scss';
import { IWpBirthdaysProps } from './IWpBirthdaysProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {  sp, List, Item, ListEnsureResult, ItemAddResult, FieldAddResult, SiteUserProps  } from '@pnp/sp';
import { SkypeForBusinessCommunicationService } from "../services";



export default class WpBirthdays extends React.Component<IWpBirthdaysProps, any> {

  constructor(props) {
    super(props);
    this.state = {
        rows: [
          { text:'?', email:'Sergey.Sokol@talansystems.onmicrosoft.com', birthday: '' }
        ],
        hash: {}
    };
  }

  public componentDidMount() {
    var self = this;
    var hash = {};
    var promises: Promise<any>[] = [];
    //const skypeService: SkypeForBusinessCommunicationService = new SkypeForBusinessCommunicationService(() => this.props.wpcontext);
    sp.web.lists.getByTitle("Employees").items
      .select("ID", "Title", "EMail", "JobTitle", "CardNumber", "Birthday", "Gender", "DepartmentRef/ID", "Contact/Id", "Contact/Title", "Contact/Name", "Contact/EMail", "Contact/JobTitle")
      .orderBy("Birthday")
      .orderBy("Title")
      .expand("DepartmentRef", "Contact")
      .get()
        .then((items: Item[]) => {
          let rows: any[] = [];
          items.map((item: Item) => {
              let parentDepRef: any = item["DepartmentRef"];
              let contact: any = item["Contact"];
              var birthday: Date = new Date(item["Birthday"]);
              var today = new Date();
              var today7 = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 7);
              today.setFullYear(2000);
              today7.setFullYear(2000);
              birthday.setFullYear(2000);
              if(birthday >= today && birthday < today7) {
                var birthdaystr : string = birthday.toLocaleDateString('ru', { month: 'long', day: 'numeric' });
                var birthdaystrarr : string[] = birthdaystr.split(" ");
                birthdaystr = birthdaystrarr[0] + ' ' + birthdaystrarr[1];              
                rows.push({ 
                  id: item["ID"], 
                  parent: (parentDepRef ? parentDepRef["ID"] : "#"), 
                  text: item["Title"], 
                  jobTitle: item["JobTitle"], 
                  email: contact["EMail"], 
                  birthday: birthdaystr  
                });
                
                //promises.push( skypeService.SubscribeToStatusChangeForUser(contact["EMail"], item["Title"], (newStatus, oldStatus, displayName) => {
                //  hash[displayName] = newStatus;
                //}));                
              }
            });
            self.setState({
              rows: rows
            });
            //Promise.all(promises).then(() => {
            //  self.setState({
            //    hash: hash
            // });        
            //});        
          })
        .catch(e => {
          alert(e); 
          self.setState({
            rows: [{text:e, email:'', birthday:''}]
          });        
        });

   
  }

  public render(): React.ReactElement<IWpBirthdaysProps> {
    return (
      <div className={ styles.wpBirthdays }>
        <div className={ styles.card }>
          <div className={ styles["card-header"] }>Іменинники на цьому тижні</div>
          <div className={ styles["card-main"] }>
               <table className={ styles.employees }>
                   { this.state.rows.map(row => 
                      <tr>
                        <td>
                          <img src={'/_vti_bin/DelveApi.ashx/people/profileimage?size=S&userId=' + row.email} alt='' height='48'/>
                        </td>
                        <td>{ row.text }</td>
                        <td>{ row.birthday }</td>
                      </tr>
                    ) }
                </table>
          </div>
        </div>      
      </div>
    );
  }

}

