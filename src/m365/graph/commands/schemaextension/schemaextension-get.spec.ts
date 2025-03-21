import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './schemaextension-get.js';

describe(commands.SCHEMAEXTENSION_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    (command as any).items = [];
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
    assert.strictEqual(command.name, commands.SCHEMAEXTENSION_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });
  it('gets schema extension', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`schemaExtensions`) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions/$entity",
          "id": "adatumisv_exo2",
          "description": "sample description",
          "targetTypes": [
            "Message"
          ],
          "status": "Available",
          "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
          "properties": [
            {
              "name": "p1",
              "type": "String"
            },
            {
              "name": "p2",
              "type": "String"
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options: {
        id: 'adatumisv_exo2'
      }
    });
    try {
      assert(loggerLogSpy.calledWith({
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions/$entity",
        "id": "adatumisv_exo2",
        "description": "sample description",
        "targetTypes": [
          "Message"
        ],
        "status": "Available",
        "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
        "properties": [
          {
            "name": "p1",
            "type": "String"
          },
          {
            "name": "p2",
            "type": "String"
          }
        ]
      }));
    }
    finally {
      sinonUtil.restore(request.get);
    }
  });
  it('gets schema extension(debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`schemaExtensions`) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions/$entity",
          "id": "adatumisv_exo2",
          "description": "sample description",
          "targetTypes": [
            "Message"
          ],
          "status": "Available",
          "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
          "properties": [
            {
              "name": "p1",
              "type": "String"
            },
            {
              "name": "p2",
              "type": "String"
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options: {
        debug: true,
        id: 'adatumisv_exo2'
      }
    });
    assert(loggerLogSpy.calledWith({
      "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions/$entity",
      "id": "adatumisv_exo2",
      "description": "sample description",
      "targetTypes": [
        "Message"
      ],
      "status": "Available",
      "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
      "properties": [
        {
          "name": "p1",
          "type": "String"
        },
        {
          "name": "p2",
          "type": "String"
        }
      ]
    }));
  });
  it('handles error', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`schemaExtensions`) > -1) {
        throw 'An error has occurred';
      }

      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        id: 'adatumisv_exo2'
      }
    } as any), new CommandError('An error has occurred'));
  });
});
