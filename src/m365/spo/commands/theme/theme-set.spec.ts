import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import { validation } from '../../../../utils/validation.js';
import commands from '../../commands.js';
import command from './theme-set.js';

describe(commands.THEME_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
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
      request.post,
      validation.isValidTheme
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.THEME_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds theme when correct parameters are passed', async () => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: 'Contoso',
        theme: '123',
        isInverted: false
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
    assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
    assert.strictEqual(postStub.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="UpdateTenantTheme" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">Contoso</Parameter><Parameter Type="String">{"isInverted":false,"name":"Contoso","palette":123}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>`);
    assert.strictEqual(loggerLogSpy.notCalled, true);
  });

  it('adds theme when correct parameters are passed (debug)', async () => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso',
        theme: '123',
        isInverted: true
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
    assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
    assert.strictEqual(postStub.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="UpdateTenantTheme" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">Contoso</Parameter><Parameter Type="String">{"isInverted":true,"name":"Contoso","palette":123}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>`);
  });

  it('handles error command error correctly', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "requestObjectIdentity ClientSvc error" } }]);
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso',
        theme: '{"isInverted":true,"name":"Contoso","palette":123}',
        inverted: false
      }
    } as any), new CommandError('requestObjectIdentity ClientSvc error'));
  });

  it('handles unknown error command error correctly', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "" } }]);
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso',
        theme: '{"isInverted":true,"name":"Contoso","palette":123}',
        inverted: false
      }
    } as any), new CommandError('ClientSvc unknown error'));
  });

  it('fails validation if the specified theme is invalid', async () => {
    const actual = await command.validate({ options: { name: 'abc', theme: '{ not valid }', isInverted: false } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when specified theme is valid', async () => {
    const theme = `{
      "themePrimary": "#d81e05",
      "themeLighterAlt": "#fdf5f4",
      "themeLighter": "#f9d6d2",
      "themeLight": "#f4b4ac",
      "themeTertiary": "#e87060",
      "themeSecondary": "#dd351e",
      "themeDarkAlt": "#c31a04",
      "themeDark": "#a51603",
      "themeDarker": "#791002",
      "neutralLighterAlt": "#eeeeee",
      "neutralLighter": "#f5f5f5",
      "neutralLight": "#e1e1e1",
      "neutralQuaternaryAlt": "#d1d1d1",
      "neutralQuaternary": "#c8c8c8",
      "neutralTertiaryAlt": "#c0c0c0",
      "neutralTertiary": "#c2c2c2",
      "neutralSecondary": "#858585",
      "neutralPrimaryAlt": "#4b4b4b",
      "neutralPrimary": "#333333",
      "neutralDark": "#272727",
      "black": "#1d1d1d",
      "white": "#f5f5f5"
    }`;
    sinon.stub(validation, 'isValidTheme').callsFake(() => true);
    const actual = await command.validate({ options: { name: 'contoso-blue', theme, isInverted: false } }, commandInfo);

    assert.strictEqual(actual, true);
  });
});
