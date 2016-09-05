import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-client-preview';

import ModuleLoader from '@microsoft/sp-module-loader';

import styles from './ToDoWebPart.module.scss';
import * as strings from 'toDoStrings';
import { IToDoWebPartWebPartProps } from './IToDoWebPartWebPartProps';

import * as angular from 'angular';
import './app/app-module';

export default class ToDoWebPartWebPart extends BaseClientSideWebPart<IToDoWebPartWebPartProps> {
  private $injector: ng.auto.IInjectorService;

  public constructor(context: IWebPartContext) {
    super(context);

    ModuleLoader.loadCss('https://appsforoffice.microsoft.com/fabric/2.6.1/fabric.min.css');
    ModuleLoader.loadCss('https://appsforoffice.microsoft.com/fabric/2.6.1/fabric.components.min.css');
  }

  public render(): void {
    if (this.renderedOnce === false) {
      this.domElement.innerHTML = `
<div class="${styles.toDoWebPart}">
  <div data-ng-controller="HomeController as vm">
    <div class="${styles.configurationNeeded}" ng-show="vm.configurationNeeded">
      Please configure the Web Part
    </div>
    <div ng-show="vm.configurationNeeded === false">
      <div class="${styles.loading}" ng-show="vm.isLoading">
        <uif-spinner>Loading...</uif-spinner>
      </div>
      <div id="entryform" ng-show="vm.isLoading === false">
        <uif-textfield uif-label="New to do:" uif-underlined ng-model="vm.newItem"
        ng-keydown="vm.todoKeyDown($event)"></uif-textfield>
      </div>
      <uif-list id="items" ng-show="vm.isLoading === false" >
        <uif-list-item ng-repeat="todo in vm.todoCollection" uif-item="todo" ng-class="{'${styles.done}': todo.done}">
          <uif-list-item-primary-text>{{todo.title}}</uif-list-item-primary-text>
          <uif-list-item-actions>
            <uif-list-item-action ng-click="vm.completeTodo(todo)" ng-show="todo.done === false">
              <uif-icon uif-type="check"></uif-icon>
            </uif-list-item-action>
            <uif-list-item-action ng-click="vm.undoTodo(todo)" ng-show="todo.done">
              <uif-icon uif-type="reactivate"></uif-icon>
            </uif-list-item-action>
            <uif-list-item-action ng-click="vm.deleteTodo(todo)">
              <uif-icon uif-type="trash"></uif-icon>
            </uif-list-item-action>
          </uif-list-item-actions>
        </uif-list-item>
      </uif-list>
    </div>
  </div>
</div>`;

      this.$injector = angular.bootstrap(this.domElement, ['todoapp']);
    }

    this.$injector.get('$rootScope').$broadcast('configurationChanged', {
      sharePointApi: this.properties.sharePointApi,
      todoListName: this.properties.todoListName,
      hideFinishedTasks: this.properties.hideFinishedTasks
    });
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.DataGroupName,
              groupFields: [
                PropertyPaneTextField('sharePointApi', {
                  label: strings.SharePointApiUrlFieldLabel
                }),
                PropertyPaneTextField('todoListName', {
                  label: strings.ToDoListNameFieldLabel
                })
              ]
            },
            {
              groupName: strings.ViewGroupName,
              groupFields: [
                PropertyPaneToggle('hideFinishedTasks', {
                  label: strings.HideFinishedTasksFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }
}
