import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './field-list.js';

describe(commands.FIELD_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FIELD_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Id', 'Title', 'InternalName', 'Hidden']);
  });

  it('fails validation if the specified site URL is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'site.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the list ID is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list id', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list title', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid and list url', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listUrl: 'sites/hr-life/Lists/breakInheritance' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if title and id are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if title and id and url are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List', listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f', listUrl: 'sites/hr-life/Lists/breakInheritance' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if title and url are specified together', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/sales', listTitle: 'Demo List', listUrl: 'sites/hr-life/Lists/breakInheritance' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('correctly handles list not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/fields`) > -1) {
        throw {
          error: {
            "odata.error": {
              "code": "-1, System.ArgumentException",
              "message": {
                "lang": "en-US",
                "value": "List 'Documents' does not exist at site with URL 'https://contoso.sharepoint.com/sites/portal'."
              }
            }
          }
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents' } } as any),
      new CommandError("List 'Documents' does not exist at site with URL 'https://contoso.sharepoint.com/sites/portal'."));
  });

  it('retrieves all site columns', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/fields`) > -1) {
        return {
          "value": [{
            "AutoIndexed": false,
            "CanBeDeleted": true,
            "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
            "ClientSideComponentProperties": null,
            "ClientValidationFormula": null,
            "ClientValidationMessage": null,
            "CustomFormatter": null,
            "DefaultFormula": null,
            "DefaultValue": null,
            "Description": "",
            "Direction": "none",
            "EnforceUniqueValues": false,
            "EntityPropertyName": "fieldname",
            "FieldTypeKind": 2,
            "Filterable": true,
            "FromBaseType": false,
            "Group": "Core Contact and Calendar Columns",
            "Hidden": false,
            "Id": "3c0e9e00-8fcc-479f-9d8d-3447cda34c5b",
            "IndexStatus": 0,
            "Indexed": false,
            "InternalName": "fieldname",
            "IsModern": false,
            "JSLink": "clienttemplates.js",
            "MaxLength": 255,
            "PinnedToFiltersPane": false,
            "ReadOnlyField": false,
            "Required": false,
            "SchemaXml": "<Field ID=\"{3C0E9E00-8FCC-479f-9D8D-3447CDA34C5B}\" Name=\"fieldname\" StaticName=\"fieldname\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" DisplayName=\"Field Name\" Group=\"Core Contact and Calendar Columns\" Type=\"Text\" Sealed=\"TRUE\" AllowDeletion=\"TRUE\" />",
            "Scope": "/sites/portal",
            "Sealed": true,
            "ShowInFiltersPane": 0,
            "Sortable": true,
            "StaticName": "fieldname",
            "Title": "Field Name",
            "TypeAsString": "Text",
            "TypeDisplayName": "Single line of text",
            "TypeShortDescription": "Single line of text",
            "ValidationFormula": null,
            "ValidationMessage": null
          }]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/portal' } });
    assert(loggerLogSpy.calledWith([
      {
        "AutoIndexed": false,
        "CanBeDeleted": true,
        "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
        "ClientSideComponentProperties": null,
        "ClientValidationFormula": null,
        "ClientValidationMessage": null,
        "CustomFormatter": null,
        "DefaultFormula": null,
        "DefaultValue": null,
        "Description": "",
        "Direction": "none",
        "EnforceUniqueValues": false,
        "EntityPropertyName": "fieldname",
        "FieldTypeKind": 2,
        "Filterable": true,
        "FromBaseType": false,
        "Group": "Core Contact and Calendar Columns",
        "Hidden": false,
        "Id": "3c0e9e00-8fcc-479f-9d8d-3447cda34c5b",
        "IndexStatus": 0,
        "Indexed": false,
        "InternalName": "fieldname",
        "IsModern": false,
        "JSLink": "clienttemplates.js",
        "MaxLength": 255,
        "PinnedToFiltersPane": false,
        "ReadOnlyField": false,
        "Required": false,
        "SchemaXml": "<Field ID=\"{3C0E9E00-8FCC-479f-9D8D-3447CDA34C5B}\" Name=\"fieldname\" StaticName=\"fieldname\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" DisplayName=\"Field Name\" Group=\"Core Contact and Calendar Columns\" Type=\"Text\" Sealed=\"TRUE\" AllowDeletion=\"TRUE\" />",
        "Scope": "/sites/portal",
        "Sealed": true,
        "ShowInFiltersPane": 0,
        "Sortable": true,
        "StaticName": "fieldname",
        "Title": "Field Name",
        "TypeAsString": "Text",
        "TypeDisplayName": "Single line of text",
        "TypeShortDescription": "Single line of text",
        "ValidationFormula": null,
        "ValidationMessage": null
      }
    ]));
  });

  it('retrieves all list columns from list queried by title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists/getByTitle('Documents')/fields`) > -1) {
        return {
          "value": [{
            "AutoIndexed": false,
            "CanBeDeleted": true,
            "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
            "ClientSideComponentProperties": null,
            "ClientValidationFormula": null,
            "ClientValidationMessage": null,
            "CustomFormatter": null,
            "DefaultFormula": null,
            "DefaultValue": null,
            "Description": "",
            "Direction": "none",
            "EnforceUniqueValues": false,
            "EntityPropertyName": "fieldname",
            "FieldTypeKind": 2,
            "Filterable": true,
            "FromBaseType": false,
            "Group": "Core Contact and Calendar Columns",
            "Hidden": false,
            "Id": "3c0e9e00-8fcc-479f-9d8d-3447cda34c5b",
            "IndexStatus": 0,
            "Indexed": false,
            "InternalName": "fieldname",
            "IsModern": false,
            "JSLink": "clienttemplates.js",
            "MaxLength": 255,
            "PinnedToFiltersPane": false,
            "ReadOnlyField": false,
            "Required": false,
            "SchemaXml": "<Field ID=\"{3C0E9E00-8FCC-479f-9D8D-3447CDA34C5B}\" Name=\"fieldname\" StaticName=\"fieldname\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" DisplayName=\"Field Name\" Group=\"Core Contact and Calendar Columns\" Type=\"Text\" Sealed=\"TRUE\" AllowDeletion=\"TRUE\" />",
            "Scope": "/sites/portal/Documents",
            "Sealed": true,
            "ShowInFiltersPane": 0,
            "Sortable": true,
            "StaticName": "fieldname",
            "Title": "Field Name",
            "TypeAsString": "Text",
            "TypeDisplayName": "Single line of text",
            "TypeShortDescription": "Single line of text",
            "ValidationFormula": null,
            "ValidationMessage": null
          }]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listTitle: 'Documents' } });
    assert(loggerLogSpy.calledWith([
      {
        "AutoIndexed": false,
        "CanBeDeleted": true,
        "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
        "ClientSideComponentProperties": null,
        "ClientValidationFormula": null,
        "ClientValidationMessage": null,
        "CustomFormatter": null,
        "DefaultFormula": null,
        "DefaultValue": null,
        "Description": "",
        "Direction": "none",
        "EnforceUniqueValues": false,
        "EntityPropertyName": "fieldname",
        "FieldTypeKind": 2,
        "Filterable": true,
        "FromBaseType": false,
        "Group": "Core Contact and Calendar Columns",
        "Hidden": false,
        "Id": "3c0e9e00-8fcc-479f-9d8d-3447cda34c5b",
        "IndexStatus": 0,
        "Indexed": false,
        "InternalName": "fieldname",
        "IsModern": false,
        "JSLink": "clienttemplates.js",
        "MaxLength": 255,
        "PinnedToFiltersPane": false,
        "ReadOnlyField": false,
        "Required": false,
        "SchemaXml": "<Field ID=\"{3C0E9E00-8FCC-479f-9D8D-3447CDA34C5B}\" Name=\"fieldname\" StaticName=\"fieldname\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" DisplayName=\"Field Name\" Group=\"Core Contact and Calendar Columns\" Type=\"Text\" Sealed=\"TRUE\" AllowDeletion=\"TRUE\" />",
        "Scope": "/sites/portal/Documents",
        "Sealed": true,
        "ShowInFiltersPane": 0,
        "Sortable": true,
        "StaticName": "fieldname",
        "Title": "Field Name",
        "TypeAsString": "Text",
        "TypeDisplayName": "Single line of text",
        "TypeShortDescription": "Single line of text",
        "ValidationFormula": null,
        "ValidationMessage": null
      }
    ]));
  });

  it('retrieves all list columns from list queried by url', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetList('%2Fsites%2Fportal%2Ftest')/fields`) > -1) {
        return {
          "value": [{
            "AutoIndexed": false,
            "CanBeDeleted": true,
            "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
            "ClientSideComponentProperties": null,
            "ClientValidationFormula": null,
            "ClientValidationMessage": null,
            "CustomFormatter": null,
            "DefaultFormula": null,
            "DefaultValue": null,
            "Description": "",
            "Direction": "none",
            "EnforceUniqueValues": false,
            "EntityPropertyName": "fieldname",
            "FieldTypeKind": 2,
            "Filterable": true,
            "FromBaseType": false,
            "Group": "Core Contact and Calendar Columns",
            "Hidden": false,
            "Id": "3c0e9e00-8fcc-479f-9d8d-3447cda34c5b",
            "IndexStatus": 0,
            "Indexed": false,
            "InternalName": "fieldname",
            "IsModern": false,
            "JSLink": "clienttemplates.js",
            "MaxLength": 255,
            "PinnedToFiltersPane": false,
            "ReadOnlyField": false,
            "Required": false,
            "SchemaXml": "<Field ID=\"{3C0E9E00-8FCC-479f-9D8D-3447CDA34C5B}\" Name=\"fieldname\" StaticName=\"fieldname\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" DisplayName=\"Field Name\" Group=\"Core Contact and Calendar Columns\" Type=\"Text\" Sealed=\"TRUE\" AllowDeletion=\"TRUE\" />",
            "Scope": "/sites/portal/Documents",
            "Sealed": true,
            "ShowInFiltersPane": 0,
            "Sortable": true,
            "StaticName": "fieldname",
            "Title": "Field Name",
            "TypeAsString": "Text",
            "TypeDisplayName": "Single line of text",
            "TypeShortDescription": "Single line of text",
            "ValidationFormula": null,
            "ValidationMessage": null
          }]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listUrl: 'test' } });
    assert(loggerLogSpy.calledWith([
      {
        "AutoIndexed": false,
        "CanBeDeleted": true,
        "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
        "ClientSideComponentProperties": null,
        "ClientValidationFormula": null,
        "ClientValidationMessage": null,
        "CustomFormatter": null,
        "DefaultFormula": null,
        "DefaultValue": null,
        "Description": "",
        "Direction": "none",
        "EnforceUniqueValues": false,
        "EntityPropertyName": "fieldname",
        "FieldTypeKind": 2,
        "Filterable": true,
        "FromBaseType": false,
        "Group": "Core Contact and Calendar Columns",
        "Hidden": false,
        "Id": "3c0e9e00-8fcc-479f-9d8d-3447cda34c5b",
        "IndexStatus": 0,
        "Indexed": false,
        "InternalName": "fieldname",
        "IsModern": false,
        "JSLink": "clienttemplates.js",
        "MaxLength": 255,
        "PinnedToFiltersPane": false,
        "ReadOnlyField": false,
        "Required": false,
        "SchemaXml": "<Field ID=\"{3C0E9E00-8FCC-479f-9D8D-3447CDA34C5B}\" Name=\"fieldname\" StaticName=\"fieldname\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" DisplayName=\"Field Name\" Group=\"Core Contact and Calendar Columns\" Type=\"Text\" Sealed=\"TRUE\" AllowDeletion=\"TRUE\" />",
        "Scope": "/sites/portal/Documents",
        "Sealed": true,
        "ShowInFiltersPane": 0,
        "Sortable": true,
        "StaticName": "fieldname",
        "Title": "Field Name",
        "TypeAsString": "Text",
        "TypeDisplayName": "Single line of text",
        "TypeShortDescription": "Single line of text",
        "ValidationFormula": null,
        "ValidationMessage": null
      }
    ]));
  });

  it('retrieves all list columns from list queried by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/web/lists(guid'3c0e9e00-8fcc-479f-9d8d-3447cda34c5b')/fields`) > -1) {
        return {
          "value": [{
            "AutoIndexed": false,
            "CanBeDeleted": true,
            "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
            "ClientSideComponentProperties": null,
            "ClientValidationFormula": null,
            "ClientValidationMessage": null,
            "CustomFormatter": null,
            "DefaultFormula": null,
            "DefaultValue": null,
            "Description": "",
            "Direction": "none",
            "EnforceUniqueValues": false,
            "EntityPropertyName": "fieldname",
            "FieldTypeKind": 2,
            "Filterable": true,
            "FromBaseType": false,
            "Group": "Core Contact and Calendar Columns",
            "Hidden": false,
            "Id": "3c0e9e00-8fcc-479f-9d8d-3447cda34c5b",
            "IndexStatus": 0,
            "Indexed": false,
            "InternalName": "fieldname",
            "IsModern": false,
            "JSLink": "clienttemplates.js",
            "MaxLength": 255,
            "PinnedToFiltersPane": false,
            "ReadOnlyField": false,
            "Required": false,
            "SchemaXml": "<Field ID=\"{3C0E9E00-8FCC-479f-9D8D-3447CDA34C5B}\" Name=\"fieldname\" StaticName=\"fieldname\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" DisplayName=\"Field Name\" Group=\"Core Contact and Calendar Columns\" Type=\"Text\" Sealed=\"TRUE\" AllowDeletion=\"TRUE\" />",
            "Scope": "/sites/portal/Documents",
            "Sealed": true,
            "ShowInFiltersPane": 0,
            "Sortable": true,
            "StaticName": "fieldname",
            "Title": "Field Name",
            "TypeAsString": "Text",
            "TypeDisplayName": "Single line of text",
            "TypeShortDescription": "Single line of text",
            "ValidationFormula": null,
            "ValidationMessage": null
          }]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/portal', listId: '3c0e9e00-8fcc-479f-9d8d-3447cda34c5b' } });
    assert(loggerLogSpy.calledWith([
      {
        "AutoIndexed": false,
        "CanBeDeleted": true,
        "ClientSideComponentId": "00000000-0000-0000-0000-000000000000",
        "ClientSideComponentProperties": null,
        "ClientValidationFormula": null,
        "ClientValidationMessage": null,
        "CustomFormatter": null,
        "DefaultFormula": null,
        "DefaultValue": null,
        "Description": "",
        "Direction": "none",
        "EnforceUniqueValues": false,
        "EntityPropertyName": "fieldname",
        "FieldTypeKind": 2,
        "Filterable": true,
        "FromBaseType": false,
        "Group": "Core Contact and Calendar Columns",
        "Hidden": false,
        "Id": "3c0e9e00-8fcc-479f-9d8d-3447cda34c5b",
        "IndexStatus": 0,
        "Indexed": false,
        "InternalName": "fieldname",
        "IsModern": false,
        "JSLink": "clienttemplates.js",
        "MaxLength": 255,
        "PinnedToFiltersPane": false,
        "ReadOnlyField": false,
        "Required": false,
        "SchemaXml": "<Field ID=\"{3C0E9E00-8FCC-479f-9D8D-3447CDA34C5B}\" Name=\"fieldname\" StaticName=\"fieldname\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" DisplayName=\"Field Name\" Group=\"Core Contact and Calendar Columns\" Type=\"Text\" Sealed=\"TRUE\" AllowDeletion=\"TRUE\" />",
        "Scope": "/sites/portal/Documents",
        "Sealed": true,
        "ShowInFiltersPane": 0,
        "Sortable": true,
        "StaticName": "fieldname",
        "Title": "Field Name",
        "TypeAsString": "Text",
        "TypeDisplayName": "Single line of text",
        "TypeShortDescription": "Single line of text",
        "ValidationFormula": null,
        "ValidationMessage": null
      }
    ]));
  });
});
