import "core-js/modules/es6.promise"; 
import "core-js/modules/es6.array.iterator.js"; 
import "core-js/modules/es6.array.from.js"; 
import "whatwg-fetch";
import "es6-map/implement";

import { Version } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import * as jQuery from "jquery";
import 'jstree';

import styles from './Wp2WebPart.module.scss';
import * as strings from 'Wp2WebPartStrings';

import {  sp, List, Item, ListEnsureResult, ItemAddResult, FieldAddResult, SiteUserProps, UserProfileQuery, SearchQuery, SearchQueryBuilder, SearchResult } from '@pnp/sp';
import { taxonomy, TermSet, TermStore, TermStores, ITermStore, ITermSetData, ITermSet, ITermData, ITerm } from "@pnp/sp-taxonomy";


import * as moment from 'moment';
//import 'moment/locale/uk';

require('jstree/dist/themes/default/style.css');
require('./Wp2.css');

export interface IWp2WebPartProps {
  description: string;
}

export interface Row {
  id: string; 
  parent: string; 
  text: string; 
  icon: string; 
  type: string;
  title: string;
  email: string;
  birthday: string;
  workphone: string;
  WorkType: string;
  TabNum: string;
  idFirm: string;
  FirmName: string;
  Auto_Card: string;
  CuratorFullName: string;
  CuratorAutoCard: string;
  MobilePhone: string;
  InternalPhone: string;
  Photo: string;
}

export default class Wp2WebPart extends BaseClientSideWebPart<IWp2WebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.wp2 }">
        <table class="panels">
          <tr>
            <td> <div id="treecontainer"> </div> </td>
            <td> <div id="panecontainer"> </div> </td> 
          </tr>            
        </table>

        <p id="message" style="color:blue;">
        </p>

        <p id="error" style="color:red;">
        </p>

      </div>`;

      //var data0x: Row[] = this.getData();

      this.getData1().then((data0: Row[]) => {

      // create tree
      var jstree: any = $('#treecontainer').jstree({ "core" : {
        "plugins": ["changed", "types"],
        "check_callback": true,
        "multiple": false,
        "data": data0
       }});

      // Get department path
      var GetDeptPath = (jstreedata: any, child_node: any, root_node_id?: string) : string => {
        var node: any = child_node;
        var texts: string[] = [];
        while(node && node.parent !== root_node_id && node.parent !== '#') {
          node = jstreedata.instance.get_node(node.parent);
          texts.unshift(node.text);
        }
        return texts.join('. ');
      };

      // Get Employee table
      var GetEmployeeTable = (jstreedata: any, person_ids: string[], root_node_id?: string, filter?: string) => {

        var ids: string[] = person_ids.sort((id1: string, id2: string) => {
          var txt1: string = jstreedata.instance.get_node(id1).text;
          var txt2: string = jstreedata.instance.get_node(id2).text;
          return txt1.localeCompare(txt2);
        });

        if(filter) {
          var exp: any = new RegExp(`${filter}`, 'i');
          ids = ids.filter((id: string) => {
            var text: string = jstreedata.instance.get_node(id).text;
            var result: boolean = (text.search(exp) >= 0);
            return result;
          });
        }

        return `<table class='table-bordered table-striped employees'>
          ${ ids.map((id: string) => { 
            var child_node: any = jstreedata.instance.get_node(id);
            if(child_node.icon === "jstree-icon jstree-file") {
              return `<tr> 
                        <td>
                          <img src='/_vti_bin/DelveApi.ashx/people/profileimage?size=S&userId=${child_node.original.email}' alt='' height='48'/>
                        </td>
                        <td><a id='${id}' class='employee-link'>${ child_node.text }</a></td>
                        <td> ${ child_node.original.title }</td>
                        <td> ${ GetDeptPath(jstreedata, child_node, root_node_id) } </td>
                        <td> ${ child_node.original.workphone }</td>
                      </tr>`;
            }
            else { 
              return ``; 
            }
          }).join(``)}
          </table>`;
      };

      jstree.on('loaded.jstree', function(e, data) {
        $(this).jstree("open_node", "root");
      });      

      jstree.on('changed.jstree', (e: any, data: any) => {
        if(data.action == "select_node") {
          var html:string = "";
          // Get node
          var node: any = data.instance.get_node(data.selected[0]);
    
          if(node.icon === "jstree-icon jstree-file") {
    
            var deps: any[] = [];
            var pers_subordinates: string[] = [];
            var deps_subordinates: string[] = [];
            var manager = "";
            data0.map(row => {
              if(row.Auto_Card == node.original.Auto_Card) {
                //var parent_node = data.instance.get_node(row.parent)
                //deps.push(parent_node.original.text);
                deps.push(GetDeptPath(data, data.instance.get_node(row.id), '56'));
              }
              if(node.original.CuratorAutoCard && row.Auto_Card == node.original.CuratorAutoCard) {
                manager = row.text;
              }
              if(row.CuratorAutoCard && row.CuratorAutoCard == node.original.Auto_Card) {
                if(row.icon === "jstree-icon jstree-file") {
                  pers_subordinates.push(row.id);
                }
                else {
                  deps_subordinates.push(row.id);
                }
              }
            });
    
            pers_subordinates = pers_subordinates.sort((id1: string, id2: string) => {
              var txt1: string = data.instance.get_node(id1).text;
              var txt2: string = data.instance.get_node(id2).text;
              return txt1.localeCompare(txt2);
            });
    
            deps_subordinates = deps_subordinates.sort((id1: string, id2: string) => {
              var txt3: string = data.instance.get_node(id1).text;
              var txt4: string = data.instance.get_node(id2).text;
              return txt3.localeCompare(txt4);
            });
    
            var deps1: any[] = [];
            var parents: string[] = node.parents;
            parents.map(parent => {
              if(parent != "#") {
                var parent_node = data.instance.get_node(parent);
                deps1.push(parent_node.original.text);
              }
            });
    
            //moment.locale('uk');
            var birthday = moment(node.original.birthday, 'DD.MM.YYYY').toDate();
            var birthdaystr : string = birthday.toLocaleDateString('ru', { month: 'long', day: 'numeric' });
            var birthdaystrarr : string[] = birthdaystr.split(" ");
            birthdaystr = birthdaystrarr[0] + ' ' + birthdaystrarr[1];

            //`<img src='/_layouts/15/userphoto.aspx?size=L&username=${node.original.email}' alt='' height='248'/>`;
    
            html = `
            <table class="panels">
              <tr>
                <td>
                  <div class="panel panel-default">
                    <div class="panel-body">
                      <img src='/_vti_bin/DelveApi.ashx/people/profileimage?size=L&userId=${node.original.email}' alt=''/>
                    </div>
                    <div class="panel-footer name-footer">
                      ${ node.original.text }
                    </div>
                  </div>          
                </td>
                <td>
                  <div class="panel panel-default">
                    <div class="panel-heading info-panel-heading"> Контактная информация </div>
                    <div class="panel-body">
                      <table class="panels">
                        <tr> <td> Рабочий телефон: </td> <td> ${ node.original.workphone } </td> </tr>
                        <tr> <td> Внутренний телефон: </td> <td>${ node.original.InternalPhone }</td> </tr>
                        <tr> <td> Мобильный телефон: </td> <td>${ node.original.MobilePhone }</td> </tr>
                        <tr> <td> Электронная почта: </td> <td> ${ node.original.email } </td> </tr>
                      </table>
                    </div>
                  </div>
    
                  <div class="panel panel-default">
                    <div class="panel-heading info-panel-heading"> Личная информация </div>
                    <div class="panel-body">
                      <table class="panels">
                        <tr> <td> День рождения: </td> <td> ${ birthdaystr } </td> </tr>
                      </table>
                    </div>
                  </div>
                </td>
              </tr>
            </table>
    
            <div class="panel panel-default">
              <div class="panel-heading info-panel-heading"> ${ data.instance.get_node(node.original.parent).text } </div>
              <div class="panel-body">
                <table class="panels">
                  <tr> <td> Должность: </td> <td> ${ node.original.title } </td> </tr>
                  <tr> <td> Подразделения: </td> <td> ${ deps.map(dep => `<span> ${dep} </span>`).join('<br/>') } </td> </tr>
                  <tr> <td> Руководитель: </td> <td> ${ manager } </td> </tr>
                </table>
              </div>
            </div>
    
            <div class="panel panel-default">
              <div class="panel-heading info-panel-heading"> Подчиненные: </div>
              <div class="panel-body">
                  <table class='table-bordered table-striped employees'>
                  ${ deps_subordinates.map( (id: string) => {
                    var child_node: any = data.instance.get_node(id);
                    return `<tr> <td> ${ GetDeptPath(data, child_node, '56') + '. ' + child_node.text } </td> </tr>`;
                  }).join('')}
                </table>
                <br/>
                ${ GetEmployeeTable(data, pers_subordinates, "56") }
              </div>
            </div>
            `;
          }
          else {
            html = `<div style="width: 400px; margin-bottom: 8px;">
                      <input id="123513451" type="search" class="form-control" placeholder="Поиск">
                    </div>
                    <div id="3454632345">
                      ${ GetEmployeeTable(data, node.children_d, node.id, "") }
                    </div>`;
          }
    
          $("#panecontainer").html(html);

          $("a.employee-link").click(function() {
            var nodex: any = data.instance.get_node($(this).attr('id'));
            data.instance.deselect_all();
            data.instance.close_all();
            data.instance.open_node(nodex);
            data.instance.select_node(nodex);
          });
    
          var search_timeout;
          $('input#123513451').on('input', () => {
            clearTimeout(search_timeout);
            search_timeout = setTimeout( () => {
              var filter: any = $('input#123513451').val();
              $('div#3454632345').html( GetEmployeeTable(data, node.children_d, node.id, filter) );
              $("a.employee-link").click(function() {
                var nodey: any = data.instance.get_node($(this).attr('id'));
                data.instance.deselect_all();
                data.instance.close_all();
                data.instance.open_node(nodey);
                data.instance.select_node(nodey);
              });
            }, 500);
          });
        }
      });
    });

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  public onInit(): Promise<void> {
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/css/bootstrap.min.css');

    sp.setup({
      spfxContext: this.context
    });

    sp.setup({
      sp: {
        headers: {
          //"Accept": "application/json; odata=nometadata"
          "Accept": "application/json; odata=minimalmetadata"
        }
      }
    });    

    return super.onInit().then(_ => {
        jQuery("#workbenchPageContent").prop("style", "max-width: none");
        jQuery(".SPCanvas-canvas").prop("style", "max-width: none");
        jQuery(".CanvasZone").prop("style", "max-width: none");
    });
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

 
  public getData1() : Promise<Row[]> {

    return new Promise<Row[]>((resolve, reject) => {
      let Rows: Row[] = [{ id: "root", parent: "#", text: "Talan Systems", icon: "jstree-icon jstree-folder", type: "department", title: "", email: "", birthday: "", workphone: "", WorkType: "", TabNum: "", idFirm: "1", FirmName: "Talan Systems", Auto_Card: "", CuratorFullName: "", CuratorAutoCard: "", MobilePhone: "", InternalPhone: "", Photo: "" }];
      //let Deps: any[] = [];
      //let Pers: any[] = [];
  
      taxonomy.getDefaultSiteCollectionTermStore()
      .groups.getByName('People')
      .termSets.getByName('Department')
      .terms.select('Id', 'Name', 'Parent', 'PathOfTerm', 'IsRoot', 'TermsCount').get()
      .then( (terms: (ITermData & ITerm)[]) => {
        terms.forEach((term: ITermData & ITerm) => {
          Rows.push({
            id: term.Name, 
            parent: (term['Parent'] ? term["Parent"].Name : 'root'), 
            text: term.Name, 
            icon: "jstree-icon jstree-folder", 
            type: "department", 
            title: "", email: "", birthday: "", workphone: "", WorkType: "", TabNum: "", idFirm: "1", FirmName: "Talan Systems", Auto_Card: "", CuratorFullName: "", CuratorAutoCard: "", MobilePhone: "", InternalPhone: "", Photo: ""
          });
        });
        //return {'Id': this.cleanGuid(term.Id), 'Name': term.Name, 'parentId': (term['Parent'] ? this.cleanGuid(term["Parent"].Id) : '#'), 'PathOfTerm': term.PathOfTerm, 'IsRoot': term.IsRoot, 'TermsCount': term.TermsCount};
  
        sp.search(<SearchQuery>{
          Querytext: '*',
          SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
          RowLimit: 1000,
          RowsPerPage: 1000,
          SelectProperties: ['AccountName', 'Department', 'JobTitle', 'WorkEmail', 'Path', 'PictureURL', 'PreferredName', 'UserProfile_GUID', 'OriginalPath']
        })
        .then(res => {
          let promises : Promise<any>[] = res.PrimarySearchResults.map((user: any) => { return sp.profiles.getPropertiesFor(user.AccountName); });
          Promise.all(promises)
          .then((allUsersProps: any[]) => {
            allUsersProps.forEach((userProps: any) => {
              let parent: any = 'root';
              userProps.UserProfileProperties.forEach((property : any) => {  
                userProps[property.Key] = property.Value;  
              });                
              if(Rows.filter((row: Row) => { return (row.id == userProps['Department']); } ).length > 0) {
                parent = userProps['Department'];
              }
              if(userProps.AccountName && userProps.Email && (userProps.Email as string).indexOf('@talansystems.onmicrosoft.com') > 0) {
                Rows.push({
                    id: userProps.AccountName, 
                    parent: parent, 
                    text: userProps.DisplayName, 
                    icon: "jstree-icon jstree-file", 
                    type: "person", 
                    title: userProps.Title, 
                    email: userProps.Email, 
                    birthday: userProps["SPS-Birthday"], 
                    workphone: userProps["WorkPhone"], 
                    WorkType: "", 
                    TabNum: "", 
                    idFirm: "1", 
                    FirmName: "Talan Systems", 
                    Auto_Card: userProps.AccountName, 
                    CuratorFullName: "", 
                    CuratorAutoCard: userProps["Manager"], 
                    MobilePhone: userProps["CellPhone"], 
                    InternalPhone: "", 
                    Photo: userProps["PictureURL"]
                });
              }


            });
            resolve(Rows);
          })
          .catch((err: any) => {
            console.log('User Properties getting error = ' + err);
            reject(err);
          });


        })
        .catch((err: any) => {
          console.log('Peoples search error = ' + err);
          reject(err);
        });

      })
      .catch((error: any) => {
        console.log('getData1 error = ' + error);
        reject(error);
      });
  
    });

    /*

    sp.search(<SearchQuery>{
      Querytext: '*',
      SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
      RowLimit: 1000,
      RowsPerPage: 1000,
      SelectProperties: ['AccountName', 'Department', 'JobTitle', 'WorkEmail', 'Path', 'PictureURL', 'PreferredName', 'UserProfile_GUID', 'OriginalPath']
    })
      .then(res => {
        console.log(res.PrimarySearchResults);
 
        Promise.all(
          res.PrimarySearchResults.map((user: any) => { sp.profiles.getPropertiesFor(user.AccountName); })
        ).then((userProps: any[]) => {


        }).catch((err: any) => {

        });
    
      })
      .catch(console.log);

    //this.setupTax()
    //.then((val: any) => {
    //  let parents = val;
    //})
    //.catch((val: any) => {
     // let error = val;
    //});

    return new Promise<Row[]>((resolve, reject) => {
      sp.web.lists.getByTitle("Departments").items.select("ID", "Title", "ParentDepartmentRef/ID").orderBy("Title").expand("ParentDepartmentRef").get().then((items: Item[]) => {
        var rows: Row[] = items.map((item: Item) => {
          let parentDepRef: any = item["ParentDepartmentRef"];
          return { id: item["ID"], parent: (parentDepRef ? parentDepRef["ID"] : "#"), text: item["Title"], icon: "jstree-icon jstree-folder", type: "department", title: "", email: "", birthday: "", workphone: "", WorkType: "", TabNum: "", idFirm: "1", FirmName: "Talan Systems", Auto_Card: "", CuratorFullName: "", CuratorAutoCard: "", MobilePhone: "", InternalPhone: "", Photo: "" };
        });

        //let user : SiteUserProps =  sp.web.siteUsers.getById(1);

        sp.web.lists.getByTitle("Employees").items.select("ID", "Title", "EMail", "JobTitle", "CardNumber", "Birthday", "Gender", "DepartmentRef/ID", "Contact/Id", "Contact/Title", "Contact/Name", "Contact/EMail").orderBy("Title").expand("DepartmentRef", "Contact").get().then((items1: Item[]) => {
          items1.map((item: Item) => {
            //let profiles: any = sp.

            let parentDepRef: any = item["DepartmentRef"];
            let contact: any = item["Contact"];
            rows.push({ id: "p" + item["ID"], parent: (parentDepRef ? parentDepRef["ID"] : "#"), text: item["Title"], icon: "jstree-icon jstree-file", type: "person", title: item["JobTitle"], email: item["EMail"], birthday: item["Birthday"], workphone: "12-34", WorkType: "", TabNum: "", idFirm: "1", FirmName: "Talan Systems", Auto_Card: item["CardNumber"], CuratorFullName: "", CuratorAutoCard: "01", MobilePhone: "03-45", InternalPhone: "", Photo: "" });
        });
          resolve(rows);
        }).catch(e => { reject(e); alert(e); });
      }).catch(e => { reject(e); });
    });
    */
    
  }

/**
   * Ensures the list exists and if it creates it adds some default example data
   */
  private ensureList(): Promise<List> {

    return new Promise<List>((resolve, reject) => {

      // use lists.ensure to always have the list available
      sp.web.lists.ensure("Departments").then((ler: ListEnsureResult) => {

        if (ler.created) {
          ler.list.get().then(list => {
            list.fields.addLookup("ParentDepartmentRef", list.Id, "Title");
          }).then(_ => {

            // and we will also add a few items so we can see some example data
            // here we use batching

            // create a batch
            let batch = sp.web.createBatch();

            ler.list.getListItemEntityTypeFullName().then(typeName => {

              ler.list.items.inBatch(batch).add({
                Title: "Title 1",
                OrderNumber: "4826492"
              }, typeName);

              ler.list.items.inBatch(batch).add({
                Title: "Title 2",
                OrderNumber: "828475"
              }, typeName);

              ler.list.items.inBatch(batch).add({
                Title: "Title 3",
                OrderNumber: "75638923"
              }, typeName);

              // excute the batched operations
              batch.execute().then(() => {
                // all of the items have been added within the batch

                resolve(ler.list);

              }).catch(e => reject(e));

            }).catch(e => reject(e));

          }).catch(e => reject(e));

        } 
        else {

          resolve(ler.list);
        }

      }).catch(e => reject(e));
    });

    
  }

  
  public async setupTax() {

    const store = await taxonomy.termStores.getByName('Taxonomy_l7ZPhzD+Gpdq56wnVOpmRA==').get();
    const termset: any = await store.getTermSetById('8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f').get();
    const terms = await termset.terms.select('Id', 'Parent', 'PathOfTerm', 'IsRoot').get();
    
    let newTerms: any = terms;
    newTerms = terms.map(term => {
      term.Id = this.cleanGuid(term.Id);
      term['PathDepth'] = term.PathOfTerm.split(';').length;
      term.TermSet = { Id: this.cleanGuid(termset.Id), Name: termset.Name };

      if (term["Parent"]) {

        term.ParentId = this.cleanGuid(term["Parent"].Id);
      }
      return term;
    });

    // RootNodes
    const parents = newTerms.filter(t => t.IsRoot);
    // Children of RootNodes
    const children = newTerms.filter(t => t.IsRoot === false);
    // Check if has children
    parents.forEach(parentTerm => {
      this._checkIfChildren(children, parentTerm);
    });

    return parents;
  }


private _checkIfChildren(items, term) {
    const children = [];
    // Loop through and check if parentId equal to parent.Id
    items.forEach(i => {
      if (i.ParentId == term.Id) {
        // Found a child, push
        children.push(i);
        // Remove this term from items and recurr
        this._checkIfChildren(items.filter(item => item !== i), i);
      }
    });

    if (children.length > 0) {
      term.Children = children;
    }
  }

public cleanGuid(guid: string): string {
    if (guid !== undefined) {
      return guid.replace('/Guid(', '').replace('/', '').replace(')', '');
    } else {
      return '';
    }
  }

}

  
