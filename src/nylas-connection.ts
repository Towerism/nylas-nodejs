import request, { UrlOptions, CoreOptions } from 'request';

import * as config from './config';

const PACKAGE_JSON = require('../package.json');
const SDK_VERSION = PACKAGE_JSON.version;
const SUPPORTED_API_VERSION = '2.1';

// The Attribute class represents a single model attribute, like 'namespace_id'
// Subclasses of Attribute like AttributeDateTime know how to covert between
// the JSON representation of that type and the javascript representation.
// The Attribute class also exposes convenience methods for generating Matchers.

export class Attribute {
  modelKey: string;
  jsonKey: string;

  constructor({ modelKey, jsonKey }: { modelKey: string; jsonKey?: string }) {
    this.modelKey = modelKey;
    this.jsonKey = jsonKey || modelKey;
  }

  toJSON(val: any) {
    return val;
  }
  fromJSON(val: any, parent: any) {
    return val || null;
  }
}

class AttributeObject extends Attribute {
  itemClass?: typeof RestfulModel;

  constructor(
    {
    modelKey,
    jsonKey,
    itemClass
  }: {
    modelKey: string;
    jsonKey?: string;
    itemClass?: typeof RestfulModel;
  }
  ) {
    super({ modelKey, jsonKey });
    this.itemClass = itemClass;
  }
}

class AttributeNumber extends Attribute {
  toJSON(val: any) {
    return val;
  }
  fromJSON(val: any, parent: any) {
    if (!isNaN(val)) {
      return Number(val);
    } else {
      return null;
    }
  }
}

class AttributeBoolean extends Attribute {
  toJSON(val: any) {
    return val;
  }
  fromJSON(val: any, parent: any) {
    return val === 'true' || val === true || false;
  }
}

class AttributeString extends Attribute {
  toJSON(val: any) {
    return val;
  }
  fromJSON(val: any, parent: any) {
    return val || '';
  }
}

class AttributeStringList extends Attribute {
  toJSON(val: any) {
    return val;
  }
  fromJSON(val: any, parent: any) {
    return val || [];
  }
}

class AttributeDate extends Attribute {
  toJSON(val: any) {
    if (!val) {
      return val;
    }
    if (!(val instanceof Date)) {
      throw new Error(
        `Attempting to toJSON AttributeDate which is not a date:
          ${this.modelKey}
         = ${val}`
      );
    }
    return val.toISOString();
  }

  fromJSON(val: any, parent: any) {
    if (!val) {
      return null;
    }
    return new Date(val);
  }
}

class AttributeDateTime extends Attribute {
  toJSON(val: any) {
    if (!val) {
      return null;
    }
    if (!(val instanceof Date)) {
      throw new Error(
        `Attempting to toJSON AttributeDateTime which is not a date:
          ${this.modelKey}
        = ${val}`
      );
    }
    return val.getTime() / 1000.0;
  }

  fromJSON(val: any, parent: any) {
    if (!val) {
      return null;
    }
    return new Date(val * 1000);
  }
}

class AttributeCollection extends Attribute {
  itemClass: typeof RestfulModel;

  constructor(
    {
    modelKey,
    jsonKey,
    itemClass
  }: {
    modelKey: string;
    jsonKey?: string;
    itemClass: typeof RestfulModel;
  }
  ) {
    super({ modelKey, jsonKey });
    this.itemClass = itemClass;
  }

  toJSON(vals: any) {
    if (!vals) {
      return [];
    }
    const json = [];
    for (const val of vals) {
      if (val.toJSON != null) {
        json.push(val.toJSON());
      } else {
        json.push(val);
      }
    }
    return json;
  }

  fromJSON(json: any, parent: any) {
    if (!json || !(json instanceof Array)) {
      return [];
    }
    const objs = [];
    for (const objJSON of json) {
      const obj = new this.itemClass(parent.connection, objJSON);
      objs.push(obj);
    }
    return objs;
  }
}


export const Attributes = {
  Number(...args: ConstructorParameters<typeof AttributeNumber>) {
    return new AttributeNumber(...args);
  },
  String(...args: ConstructorParameters<typeof AttributeString>) {
    return new AttributeString(...args);
  },
  StringList(...args: ConstructorParameters<typeof AttributeStringList>) {
    return new AttributeStringList(...args);
  },
  DateTime(...args: ConstructorParameters<typeof AttributeDateTime>) {
    return new AttributeDateTime(...args);
  },
  Date(...args: ConstructorParameters<typeof AttributeDate>) {
    return new AttributeDate(...args);
  },
  Collection(...args: ConstructorParameters<typeof AttributeCollection>) {
    return new AttributeCollection(...args);
  },
  Boolean(...args: ConstructorParameters<typeof AttributeBoolean>) {
    return new AttributeBoolean(...args);
  },
  Object(...args: ConstructorParameters<typeof AttributeObject>) {
    return new AttributeObject(...args);
  }
};

const REQUEST_CHUNK_SIZE = 100;

export type GetCallback = (error: Error | null, result?: RestfulModel) => void;

export class RestfulModelCollection<T extends RestfulModel> {
  connection: NylasConnection;
  modelClass: typeof RestfulModel;

  constructor(modelClass: typeof RestfulModel, connection: NylasConnection) {
    this.modelClass = modelClass;
    this.connection = connection;
    if (!(this.connection instanceof NylasConnection)) {
      throw new Error('Connection object not provided');
    }
    if (!this.modelClass) {
      throw new Error('Model class not provided');
    }
  }

  forEach(
    params: { [key: string]: any } = {},
    eachCallback: (item: T) => void,
    completeCallback?: (err?: Error | null | undefined) => void
  ) {
    if (params.view == 'count') {
      const err = new Error('forEach() cannot be called with the count view');
      if (completeCallback) {
        completeCallback(err);
      }
      return Promise.reject(err);
    }

    let offset = 0;

    const iteratee = (): Promise<void> => {
      return this._getItems(params, offset, REQUEST_CHUNK_SIZE).then(items => {
        for (const item of items) {
          eachCallback(item);
        }
        offset += items.length;
        const finished = items.length < REQUEST_CHUNK_SIZE;
        if (finished === false) {
          return iteratee();
        }
      });
    };

    iteratee().then(
      () => {
        if (completeCallback) {
          completeCallback();
        }
      },
      (err: Error) => {
        if (completeCallback) {
          return completeCallback(err);
        }
      }
    );
  }

  count(
    params: { [key: string]: any } = {},
    callback?: (err: Error | null, num?: number) => void
  ) {
    return this.connection
      .request({
        method: 'GET',
        path: this.path(),
        qs: { view: 'count', ...params },
      })
      .then((json: any) => {
        if (callback) {
          callback(null, json.count);
        }
        return Promise.resolve(json.count);
      })
      .catch((err: Error) => {
        if (callback) {
          callback(err);
        }
        return Promise.reject(err);
      });
  }

  first(
    params: { [key: string]: any } = {},
    callback?: (error: Error | null, model?: T) => void
  ) {
    if (params.view == 'count') {
      const err = new Error('first() cannot be called with the count view');
      if (callback) {
        callback(err);
      }
      return Promise.reject(err);
    }

    return this._getItems(params, 0, 1)
      .then(items => {
        if (callback) {
          callback(null, items[0]);
        }
        return Promise.resolve(items[0]);
      })
      .catch(err => {
        if (callback) {
          callback(err);
        }
        return Promise.reject(err);
      });
  }

  list(
    params: { [key: string]: any } = {},
    callback?: (error: Error | null, obj?: T[]) => void
  ) {
    if (params.view == 'count') {
      const err = new Error('list() cannot be called with the count view');
      if (callback) {
        callback(err);
      }
      return Promise.reject(err);
    }

    const limit = params.limit || Infinity;
    const offset = params.offset;
    return this._range({ params, offset, limit, callback });
  }

  find(
    id: string,
    paramsArg?: { [key: string]: any } | GetCallback | null,
    callbackArg?: GetCallback | { [key: string]: any } | null
  ) {

    // callback used to be the second argument, and params was the third
    let callback: GetCallback | undefined;
    if (typeof callbackArg === 'function') {
      callback = callbackArg as GetCallback;
    } else if (typeof paramsArg === 'function') {
      callback = paramsArg as GetCallback;
    }

    let params: { [key: string]: any } = {};
    if (paramsArg && typeof paramsArg === 'object') {
      params = paramsArg;
    } else if (callbackArg && typeof callbackArg === 'object') {
      params = callbackArg;
    }

    if (!id) {
      const err = new Error('find() must be called with an item id');
      if (callback) {
        callback(err);
      }
      return Promise.reject(err);
    }

    if (params.view == 'count' || params.view == 'ids') {
      const err = new Error(
        'find() cannot be called with the count or ids view'
      );
      if (callback) {
        callback(err);
      }
      return Promise.reject(err);
    }

    return this._getModel(id, params)
      .then(model => {
        if (callback) {
          callback(null, model);
        }
        return Promise.resolve(model);
      })
      .catch(err => {
        if (callback) {
          callback(err);
        }
        return Promise.reject(err);
      });
  }

  delete(
    itemOrId: T | string,
    params: { [key: string]: any } = {},
    callback?: (error: Error | null) => void
  ) {
    if (!itemOrId) {
      const err = new Error('delete() requires an item or an id');
      if (callback) {
        callback(err);
      }
      return Promise.reject(err);
    }

    if (typeof params === 'function') {
      callback = params as (error: Error | null) => void;
      params = {};
    }

    const item =
      typeof itemOrId === 'string' ? this.build({ id: itemOrId }) : itemOrId;

    const options: { [key: string]: any } = item.deleteRequestOptions(params);
    options.item = item;

    return this.deleteItem(options, callback);
  }

  deleteItem(
    options: { [key: string]: any },
    callbackArg?: (error: Error | null) => void
  ) {
    const item = options.item;
    // callback used to be in the options object
    const callback = options.callback ? options.callback : callbackArg;
    const body = options.hasOwnProperty('body')
      ? options.body
      : item.deleteRequestBody({});
    const qs = options.hasOwnProperty('qs')
      ? options.qs
      : item.deleteRequestQueryString({});

    return this.connection
      .request({
        method: 'DELETE',
        qs: qs,
        body: body,
        path: `${this.path()}/${item.id}`,
      })
      .then((data) => {
        if (callback) {
          callback(null, data);
        }
        return Promise.resolve(data);
      })
      .catch((err: Error) => {
        if (callback) {
          callback(err);
        }
        return Promise.reject(err);
      });
  }

  build(args: { [key: string]: any }) {
    const model = this._createModel({});
    for (const key in args) {
      const val = args[key];
      (model as any)[key] = val;
    }
    return model;
  }

  path() {
    return `/${this.modelClass.collectionName}`;
  }

  _range({
    params = {},
    offset = 0,
    limit = 100,
    callback,
    path,
  }: {
    params?: { [key: string]: any };
    offset?: number;
    limit?: number;
    callback?: (error: Error | null, results?: T[]) => void;
    path?: string;
  }) {
    let accumulated: T[] = [];

    const iteratee = (): Promise<void> => {
      const chunkOffset = offset + accumulated.length;
      const chunkLimit = Math.min(
        REQUEST_CHUNK_SIZE,
        limit - accumulated.length
      );
      return this._getItems(params, chunkOffset, chunkLimit, path).then(
        items => {
          accumulated = accumulated.concat(items);
          const finished =
            items.length < REQUEST_CHUNK_SIZE || accumulated.length >= limit;
          if (finished === false) {
            return iteratee();
          }
        }
      );
    };

    // do not return rejected promise when callback is provided
    // to prevent unhandled rejection warning
    return iteratee().then(
      () => {
        if (callback) {
          return callback(null, accumulated);
        }
        return accumulated;
      },
      (err: Error) => {
        if (callback) {
          return callback(err);
        }
        throw err;
      }
    );
  }

  _getItems(
    params: { [key: string]: any },
    offset: number,
    limit: number,
    path?: string
  ): Promise<T[]> {
    // Items can be either models or ids

    if (!path) {
      path = this.path();
    }

    if (params.view == 'ids') {
      return this.connection.request({
        method: 'GET',
        path,
        qs: { ...params, offset, limit },
      });
    }

    return this._getModelCollection(params, offset, limit, path);
  }

  _createModel(json: { [key: string]: any }) {
    return new this.modelClass(this.connection, json) as T;
  }

  _getModel(id: string, params: { [key: string]: any } = {}): Promise<T> {
    return this.connection
      .request({
        method: 'GET',
        path: `${this.path()}/${id}`,
        qs: params,
      })
      .then((json: any) => {
        const model = this._createModel(json);
        return Promise.resolve(model);
      });
  }

  _getModelCollection(
    params: { [key: string]: any },
    offset?: number,
    limit?: number,
    path?: string
  ): Promise<T[]> {
    return this.connection
      .request({
        method: 'GET',
        path,
        qs: { ...params, offset, limit },
      })
      .then((jsonArray: any) => {
        const models = jsonArray.map((json: any) => {
          return this._createModel(json);
        });
        return Promise.resolve(models);
      });
  }
}

export type SaveCallback = (error: Error | null, result?: RestfulModel) => void;

interface RestfulModelJSON {
  id: string;
  object: string;
  accountId: string;
  [key: string]: any;
}

export class RestfulModel {
  static endpointName = ''; // overrridden in subclasses
  static collectionName = ''; // overrridden in subclasses
  static attributes: { [key: string]: Attribute };

  accountId?: string;
  connection: NylasConnection;
  id?: string;
  object?: string;

  constructor(connection: NylasConnection, json?: Partial<RestfulModelJSON>) {
    this.connection = connection;
    if (!(this.connection instanceof NylasConnection)) {
      throw new Error('Connection object not provided');
    }
    if (json) {
      this.fromJSON(json);
    }
  }

  attributes(): { [key: string]: Attribute } {
    return (this.constructor as any).attributes;
  }

  isEqual(other: RestfulModel) {
    return (
      (other ? other.id : undefined) === this.id &&
      (other ? other.constructor : undefined) === this.constructor
    );
  }

  fromJSON(json: Partial<RestfulModelJSON> = {}) {
    const attributes = this.attributes();
    for (const attrName in attributes) {
      const attr = attributes[attrName];
      if (json[attr.jsonKey] !== undefined) {
        (this as any)[attrName] = attr.fromJSON(json[attr.jsonKey], this);
      }
    }
    return this;
  }

  toJSON() {
    const json: any = {};
    const attributes = this.attributes();
    for (const attrName in attributes) {
      const attr = attributes[attrName];
      json[attr.jsonKey] = attr.toJSON((this as any)[attrName]);
    }
    json['object'] = this.constructor.name.toLowerCase();
    return json;
  }

  // Subclasses should override this method.
  pathPrefix() {
    return '';
  }

  saveEndpoint() {
    const collectionName = (this.constructor as any).collectionName;
    return `${this.pathPrefix()}/${collectionName}`;
  }

  // saveRequestBody is used by save(). It returns a JSON dict containing only the
  // fields the API allows updating. Subclasses should override this method.
  saveRequestBody() {
    return this.toJSON();
  }

  // deleteRequestQueryString is used by delete(). Subclasses should override this method.
  deleteRequestQueryString(params: { [key: string]: any }) {
    return {};
  }
  // deleteRequestBody is used by delete(). Subclasses should override this method.
  deleteRequestBody(params: { [key: string]: any }) {
    return {};
  }

  // deleteRequestOptions is used by delete(). Subclasses should override this method.
  deleteRequestOptions(params: { [key: string]: any }) {
    return {
      body: this.deleteRequestBody(params),
      qs: this.deleteRequestQueryString(params),
    };
  }

  toString() {
    return JSON.stringify(this.toJSON());
  }

  // Not every model needs to have a save function, but those who
  // do shouldn't have to reimplement the same boilerplate.
  // They should instead define a save() function which calls _save.
  _save(params: {} | SaveCallback = {}, callback?: SaveCallback) {
    if (typeof params === 'function') {
      callback = params as SaveCallback;
      params = {};
    }
    return this.connection
      .request({
        method: this.id ? 'PUT' : 'POST',
        body: this.saveRequestBody(),
        qs: params,
        path: this.id
          ? `${this.saveEndpoint()}/${this.id}`
          : `${this.saveEndpoint()}`,
      })
      .then(json => {
        this.fromJSON(json as RestfulModelJSON);
        if (callback) {
          callback(null, this);
        }
        return Promise.resolve(this);
      })
      .catch(err => {
        if (callback) {
          callback(err);
        }
        return Promise.reject(err);
      });
  }

  _get(
    params: { [key: string]: any } = {},
    callback?: (error: Error | null, result?: any) => void,
    path_suffix = ''
  ) {
    const collectionName = (this.constructor as any).collectionName;
    return this.connection
      .request({
        method: 'GET',
        path: `/${collectionName}/${this.id}${path_suffix}`,
        qs: params,
      })
      .then(response => {
        if (callback) {
          callback(null, response);
        }
        return Promise.resolve(response);
      })
      .catch(err => {
        if (callback) {
          callback(err);
        }
        return Promise.reject(err);
      });
  }
}
(RestfulModel as any).attributes = {
  id: Attributes.String({
    modelKey: 'id',
  }),
  object: Attributes.String({
    modelKey: 'object',
  }),
  accountId: Attributes.String({
    modelKey: 'accountId',
    jsonKey: 'account_id',
  }),
};

export class RestfulModelInstance {
  connection: NylasConnection;
  modelClass: typeof RestfulModel;

  constructor(modelClass: typeof RestfulModel, connection: NylasConnection) {
    this.modelClass = modelClass;
    this.connection = connection;
    if (!(this.connection instanceof NylasConnection)) {
      throw new Error('Connection object not provided');
    }
    if (!this.modelClass) {
      throw new Error('Model class not provided');
    }
  }

  path() {
    return `/${this.modelClass.endpointName}`;
  }

  get(params: { [key: string]: any } = {}) {
    return this.connection
      .request({
        method: 'GET',
        path: this.path(),
        qs: params,
      })
      .then(json => {
        const model = new this.modelClass(this.connection, json);
        return Promise.resolve(model);
      });
  }
}

export class Calendar extends RestfulModel {
  name?: string;
  description?: string;
  readOnly?: boolean;
  location?: string;
  timezone?: string;
  isPrimary?: boolean;
  jobStatusId?: string;

  save(params: {} | SaveCallback = {}, callback?: SaveCallback) {
    return this._save(params, callback);
  }

  saveRequestBody() {
    const calendarJSON = this.toJSON();
    return {
      name: calendarJSON.name,
      description: calendarJSON.description,
      location: calendarJSON.location,
      timezone: calendarJSON.timezone
    };
  }
}

Calendar.collectionName = 'calendars';
Calendar.attributes = {
  ...RestfulModel.attributes,
  name: Attributes.String({
    modelKey: 'name',
  }),
  description: Attributes.String({
    modelKey: 'description',
  }),
  readOnly: Attributes.Boolean({
    modelKey: 'readOnly',
    jsonKey: 'read_only',
  }),
  location: Attributes.String({
    modelKey: 'location',
  }),
  timezone: Attributes.String({
    modelKey: 'timezone',
  }),
  isPrimary: Attributes.Boolean({
    modelKey: 'isPrimary',
    jsonKey: 'is_primary',
  }),
  jobStatusId: Attributes.String({
    modelKey: 'jobStatusId',
    jsonKey: 'job_status_id'
  })
};

export class CalendarRestfulModelCollection<Calendar> extends RestfulModelCollection<RestfulModel> {
  connection: NylasConnection;
  modelClass: typeof Calendar;

  constructor(connection: NylasConnection) {
    super(Calendar, connection);
    this.connection = connection;
    this.modelClass = Calendar;
  }

  freeBusy(options: {
    start_time?: string,
    startTime?: string,
    end_time?: string,
    endTime?: string,
    emails: string[]
  }, callback?: (error: Error | null, data?: { [key: string]: any }) => void) {

    return this.connection
      .request({
        method: 'POST',
        path: `/calendars/free-busy`,
        body: {
          start_time: options.startTime || options.start_time,
          end_time: options.endTime || options.end_time,
          emails: options.emails
        }
      })
      .then(json => {
        if (callback) {
          callback(null, json);
        }
        return Promise.resolve(json);
      })
      .catch(err => {
        if (callback) {
          callback(err);
        }
        return Promise.reject(err);
      });
  }
}

export class EventParticipant extends RestfulModel {
  name?: string;
  email?: string;
  status?: string;

  toJSON() {
    const json = super.toJSON();
    if (!json['name']) {
      json['name'] = json['email'];
    }
    delete json['object'];
    return json;
  }
}
EventParticipant.collectionName = 'event-participants';
EventParticipant.attributes = {
  name: Attributes.String({
    modelKey: 'name',
  }),
  email: Attributes.String({
    modelKey: 'email',
  }),
  status: Attributes.String({
    modelKey: 'status',
  }),
};

export class Event extends RestfulModel {
  calendarId?: string;
  iCalUID?: string;
  messageId?: string;
  title?: string;
  description?: string;
  owner?: string;
  participants?: EventParticipant[];
  readOnly?: boolean;
  location?: string;
  masterEventId?: string;
  when?: {
    start_time?: number;
    end_time?: number;
    time?: number;
    start_date?: string;
    end_date?: string;
    date?: string;
    object?: string;
  };
  busy?: boolean;
  status?: string;
  recurrence?: {
    rrule: string[];
    timezone: string;
  };

  get start() {
    const start =
      this.when?.start_time ||
      this.when?.start_date ||
      this.when?.time ||
      this.when?.date;
    return start;
  }

  set start(val: string | number | undefined) {
    if (!this.when) {
      this.when = {};
    }
    if (typeof val === 'number') {
      if (val === this.when.end_time) {
        this.when = { time: val };
      } else {
        delete this.when.time;
        delete this.when.start_date;
        delete this.when.date;
        this.when.start_time = val;
      }
    }
    if (typeof val === 'string') {
      if (val === this.when.end_date) {
        this.when = { date: val };
      } else {
        delete this.when.date;
        delete this.when.start_time;
        delete this.when.time;
        this.when.start_date = val;
      }
    }
  }

  get end() {
    const end =
      this.when?.end_time ||
      this.when?.end_date ||
      this.when?.time ||
      this.when?.date;
    return end;
  }

  set end(val: string | number | undefined) {
    if (!this.when) {
      this.when = {};
    }
    if (typeof val === 'number') {
      if (val === this.when.start_time) {
        this.when = { time: val };
      } else {
        delete this.when.time;
        delete this.when.end_date;
        delete this.when.date;
        this.when.end_time = val;
      }
    }
    if (typeof val === 'string') {
      if (val === this.when.start_date) {
        this.when = { date: val };
      } else {
        delete this.when.date;
        delete this.when.time;
        delete this.when.end_time;
        this.when.end_date = val;
      }
    }
  }

  deleteRequestQueryString(params: { [key: string]: any } = {}) {
    const qs: { [key: string]: any } = {};
    if (params.hasOwnProperty('notify_participants')) {
      qs.notify_participants = params.notify_participants;
    }
    return qs;
  }

  save(params: {} | SaveCallback = {}, callback?: SaveCallback) {
    return this._save(params, callback);
  }

  rsvp(status: string, comment: string, callback: (error: Error | null, data?: Event) => void) {
    return this.connection
      .request({
        method: 'POST',
        body: { event_id: this.id, status: status, comment: comment },
        path: '/send-rsvp',
      })
      .then(json => {
        this.fromJSON(json);
        if (callback) {
          callback(null, this);
        }
        return Promise.resolve(this);
      })
      .catch(err => {
        if (callback) {
          callback(err);
        }
        return Promise.reject(err);
      });
  }
}
Event.collectionName = 'events';
Event.attributes = {
  ...RestfulModel.attributes,
  calendarId: Attributes.String({
    modelKey: 'calendarId',
    jsonKey: 'calendar_id',
  }),
  masterEventId: Attributes.String({
    modelKey: 'masterEventId',
    jsonKey: 'master_event_id',
  }),
  iCalUID: Attributes.String({
    modelKey: 'iCalUID',
    jsonKey: 'ical_uid',
  }),
  messageId: Attributes.String({
    modelKey: 'messageId',
    jsonKey: 'message_id',
  }),
  title: Attributes.String({
    modelKey: 'title',
  }),
  description: Attributes.String({
    modelKey: 'description',
  }),
  owner: Attributes.String({
    modelKey: 'owner',
  }),
  participants: Attributes.Collection({
    modelKey: 'participants',
    itemClass: EventParticipant,
  }),
  readOnly: Attributes.Boolean({
    modelKey: 'readOnly',
    jsonKey: 'read_only',
  }),
  location: Attributes.String({
    modelKey: 'location',
  }),
  when: Attributes.Object({
    modelKey: 'when',
  }),
  busy: Attributes.Boolean({
    modelKey: 'busy',
  }),
  status: Attributes.String({
    modelKey: 'status',
  }),
  recurrence: Attributes.Object({
    modelKey: 'recurrence',
  })
};

export class Account extends RestfulModel {
  name?: string;
  emailAddress?: string;
  provider?: string;
  organizationUnit?: string;
  syncState?: string;
  billingState?: string;
  linkedAt?: Date;
}
Account.collectionName = 'accounts';
Account.endpointName = 'account';
Account.attributes = {
  ...RestfulModel.attributes,
  name: Attributes.String({
    modelKey: 'name',
  }),

  emailAddress: Attributes.String({
    modelKey: 'emailAddress',
    jsonKey: 'email_address',
  }),

  provider: Attributes.String({
    modelKey: 'provider',
  }),

  organizationUnit: Attributes.String({
    modelKey: 'organizationUnit',
    jsonKey: 'organization_unit',
  }),

  syncState: Attributes.String({
    modelKey: 'syncState',
    jsonKey: 'sync_state',
  }),

  billingState: Attributes.String({
    modelKey: 'billingState',
    jsonKey: 'billing_state',
  }),

  linkedAt: Attributes.DateTime({
    modelKey: 'linkedAt',
    jsonKey: 'linked_at',
  }),
};

export default class NylasConnection {
  accessToken: string | null | undefined;
  clientId: string | null | undefined;

  calendars: CalendarRestfulModelCollection<Calendar> = new CalendarRestfulModelCollection(this);
  events: RestfulModelCollection<Event> = new RestfulModelCollection(Event, this);
  account = new RestfulModelInstance(Account, this);

  constructor(
    accessToken: string | null | undefined,
    { clientId }: { clientId: string | null | undefined }
  ) {
    this.accessToken = accessToken;
    this.clientId = clientId;
  }

  requestOptions(options: { [key: string]: any }) {
    if (!options) {
      options = {};
    }
    options = { ...options };
    if (!options.method) {
      options.method = 'GET';
    }
    if (options.path) {
      if (!options.url) {
        options.url = `${config.apiServer}${options.path}`;
      }
    }
    if (!options.formData) {
      if (!options.body) {
        options.body = {};
      }
    }
    if (options.json == null) {
      options.json = true;
    }
    if (!options.downloadRequest) {
      options.downloadRequest = false;
    }

    // For convenience, If `expanded` param is provided, convert to view:
    // 'expanded' api option
    if (options.qs && options.qs.expanded) {
      if (options.qs.expanded === true) {
        options.qs.view = 'expanded';
      }
      delete options.qs.expanded;
    }

    const user =
      options.path.substr(0, 3) === '/a/'
        ? config.clientSecret
        : this.accessToken;

    if (user) {
      options.auth = {
        user: user,
        pass: '',
        sendImmediately: true,
      };
    }

    if (options.headers == null) {
      options.headers = {};
    }
    if (options.headers['User-Agent'] == null) {
      options.headers['User-Agent'] = `Nylas Node SDK v${SDK_VERSION}`;
    }

    options.headers['Nylas-SDK-API-Version'] = SUPPORTED_API_VERSION;
    options.headers['X-Nylas-Client-Id'] = this.clientId;

    return options as (CoreOptions & UrlOptions & { downloadRequest: boolean });
  }
  _getWarningForVersion(sdkApiVersion?: string, apiVersion?: string) {
    let warning = '';

    if (sdkApiVersion != apiVersion) {
      if (sdkApiVersion && apiVersion) {
        warning +=
          `WARNING: SDK version may not support your Nylas API version.` +
          ` SDK supports version ${sdkApiVersion} of the API and your application` +
          ` is currently running on version ${apiVersion} of the API.`;

        const apiNum = parseInt(apiVersion.split('-')[0]);
        const sdkNum = parseInt(sdkApiVersion.split('-')[0]);

        if (sdkNum > apiNum) {
          warning += ` Please update the version of the API that your application is using through the developer dashboard.`;
        } else if (apiNum > sdkNum) {
          warning += ` Please update the sdk to ensure it works properly.`;
        }
      }
    }
    return warning;
  }
  request(options?: Parameters<this['requestOptions']>[0]) {
    if (!options) {
      options = {};
    }
    const resolvedOptions = this.requestOptions(options);

    return new Promise<any>((resolve, reject) => {
      return request(resolvedOptions, (error, response, body = {}) => {
        if (typeof response === 'undefined') {
          error = error || new Error('No response');
          return reject(error);
        }
        // node headers are lowercase so this refers to `Nylas-Api-Version`
        const apiVersion = response.headers['nylas-api-version'] as
          | string
          | undefined;

        const warning = this._getWarningForVersion(
          SUPPORTED_API_VERSION,
          apiVersion
        );
        if (warning) {
          console.warn(warning);
        }

        // raw MIMI emails have json === false and the body is a string so
        // we need to turn into JSON before we can access fields
        if (resolvedOptions.json === false) {
          body = JSON.parse(body);
        }

        if (error || response.statusCode > 299) {
          if (!error) {
            error = new Error(body.message);
          }
          if (body.missing_fields) {
            error.message = `${body.message}: ${body.missing_fields}`;
          }
          if (body.server_error) {
            error.message = `${error.message} (Server Error:
              ${body.server_error}
            )`;
          }
          if (response.statusCode) {
            error.statusCode = response.statusCode;
          }
          return reject(error);
        } else {
          if (resolvedOptions.downloadRequest) {
            return resolve(response);
          } else {
            return resolve(body);
          }
        }
      });
    });
  }
}
