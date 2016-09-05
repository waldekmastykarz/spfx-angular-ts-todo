export interface ITodo {
  id: number;
  title: string;
  done: boolean;
}

interface ITodoItem {
  Id: number;
  Title: string;
  Status: string;
}

export interface IDataService {
  getTodos(sharePointApi: string, todoListName: string, hideFinishedTasks: boolean): ng.IPromise<ITodo[]>;
  addTodo(todo: string, sharePointApi: string, todoListName: string): ng.IPromise<{}>;
  deleteTodo(todo: ITodo, sharePointApi: string, todoListName: string): ng.IPromise<{}>;
  setTodoStatus(todo: ITodo, done: boolean, sharePointApi: string, todoListName: string): ng.IPromise<{}>;
}

export default class DataService implements IDataService {
  public static $inject: string[] = ['$q', '$http'];

  constructor(private $q: ng.IQService, private $http: ng.IHttpService) {
  }

  public getTodos(sharePointApi: string, todoListName: string, hideFinishedTasks: boolean): ng.IPromise<ITodo[]> {
    const deferred: ng.IDeferred<ITodo[]> = this.$q.defer();

    let url: string =
      `${sharePointApi}web/lists/getbytitle('${todoListName}')/items?$select=Id,Title,Status&$orderby=ID desc`;

    if (hideFinishedTasks === true) {
      url += "&$filter=Status ne 'Completed'";
    }

    this.$http({
      url: url,
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=nometadata'
      }
    }).then((result: ng.IHttpPromiseCallbackArg<{ value: ITodoItem[] }>): void => {
      const todos: ITodo[] = [];
      for (let i: number = 0; i < result.data.value.length; i++) {
        const todo: ITodoItem = result.data.value[i];
        todos.push({
          id: todo.Id,
          title: todo.Title,
          done: todo.Status === 'Completed'
        });
      }
      deferred.resolve(todos);
    });

    return deferred.promise;
  }

  public addTodo(todo: string, sharePointApi: string, todoListName: string): ng.IPromise<{}> {
    const deferred: ng.IDeferred<{}> = this.$q.defer();

    this.$http({
      url: sharePointApi + 'contextinfo',
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=nometadata'
      }
    }).then((digestResult: ng.IHttpPromiseCallbackArg<{ FormDigestValue: string }>): void => {
      const requestDigest: string = digestResult.data.FormDigestValue;
      const body: string = JSON.stringify({
        '__metadata': {
          'type': 'SP.Data.' +
            todoListName.charAt(0).toUpperCase() +
            todoListName.slice(1) + 'ListItem'
          },
        'Title': todo
      });
      this.$http({
        url: sharePointApi + 'web/lists/getbytitle(\'' + todoListName + '\')/items',
        method: 'POST',
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'X-RequestDigest': requestDigest
        },
        data: body
      }).then((result: ng.IHttpPromiseCallbackArg<{}>): void => {
        deferred.resolve();
      });
    });

    return deferred.promise;
  }

  public deleteTodo(todo: ITodo, sharePointApi: string, todoListName: string): ng.IPromise<{}> {
    const deferred: ng.IDeferred<{}> = this.$q.defer();

    this.$http({
      url: sharePointApi + 'contextinfo',
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=nometadata'
      }
    }).then((digestResult: ng.IHttpPromiseCallbackArg<{ FormDigestValue: string }>): void => {
      const requestDigest: string = digestResult.data.FormDigestValue;
      this.$http({
        url: sharePointApi + 'web/lists/getbytitle(\'' + todoListName + '\')/items(' + todo.id + ')',
        method: 'POST',
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'X-RequestDigest': requestDigest,
          'IF-MATCH': '*',
          'X-HTTP-Method': 'DELETE'
        }
      }).then((result: ng.IHttpPromiseCallbackArg<{}>): void => {
        deferred.resolve();
      });
    });

    return deferred.promise;
  }

  public setTodoStatus(todo: ITodo, done: boolean, sharePointApi: string, todoListName: string): ng.IPromise<{}> {
    const deferred: ng.IDeferred<{}> = this.$q.defer();

    this.$http({
      url: sharePointApi + 'contextinfo',
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=nometadata'
      }
    }).then((digestResult: ng.IHttpPromiseCallbackArg<{ FormDigestValue: string }>): void => {
      const requestDigest: string = digestResult.data.FormDigestValue;
      const body: string = JSON.stringify({
        '__metadata': {
          'type': 'SP.Data.' +
            todoListName.charAt(0).toUpperCase() +
            todoListName.slice(1) + 'ListItem'
          },
        'Status': done ? 'Completed' : 'Not started'
      });
      this.$http({
        url: sharePointApi + 'web/lists/getbytitle(\'' + todoListName + '\')/items(' + todo.id + ')',
        method: 'POST',
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'X-RequestDigest': requestDigest,
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        data: body
      }).then((result: ng.IHttpPromiseCallbackArg<{}>): void => {
        deferred.resolve();
      });
    });

    return deferred.promise;
  }
}
