import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './bucket-get.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.BUCKET_GET, () => {
  const validBucketId = 'vncYUXCRBke28qMLB-d4xJcACtNz';
  const validBucketName = 'Bucket name';
  const validPlanId = 'oUHpnKBFekqfGE_PS6GGUZcAFY7b';
  const validPlanTitle = 'Plan title';
  const validOwnerGroupName = 'Group name';
  const validOwnerGroupId = '00000000-0000-0000-0000-000000000000';
  const invalidOwnerGroupId = 'Invalid GUID';
  const validRosterId = 'RuY-PSpdw02drevnYDTCJpgAEfoI';

  const singleGroupResponse = {
    "value": [
      {
        "id": validOwnerGroupId,
        "displayName": validOwnerGroupName
      }
    ]
  };

  const multipleGroupResponse = {
    "value": [
      {
        "id": validOwnerGroupId
      },
      {
        "id": validOwnerGroupId
      }
    ]
  };

  const singlePlanResponse = {
    "value": [
      {
        "id": validPlanId,
        "title": validPlanTitle
      }
    ]
  };

  const singleBucketByNameResponse = {
    "value": [
      {
        "name": validBucketName,
        "id": validBucketId
      }
    ]
  };

  const singleBucketByIdResponse = {
    "name": validBucketName,
    "id": validBucketId
  };

  const multipleBucketByNameResponse = {
    "value": [
      {
        "name": validBucketName,
        "id": validBucketId
      },
      {
        "name": validBucketName,
        "id": validBucketId
      }
    ]
  };

  const planResponse = {
    value: [{
      id: validPlanId,
      title: validPlanTitle
    }]
  };

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
    auth.connection.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
    commandInfo = cli.getCommandInfo(command);
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === 'prompt') {
        return false;
      }

      return defaultValue;
    });
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
      request.get,
      request.patch,
      cli.handleMultipleResultsFound,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.BUCKET_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when no option is specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when id and plan details are specified', async () => {
    const actual = await command.validate({
      options: {
        id: validBucketId,
        planId: validPlanId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when owner group id is not a guid', async () => {
    const actual = await command.validate({
      options: {
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupId: invalidOwnerGroupId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when plan id is used with owner group name', async () => {
    const actual = await command.validate({
      options: {
        name: validBucketName,
        planId: validPlanId,
        ownerGroupName: validOwnerGroupName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when plan id is used with owner group id', async () => {
    const actual = await command.validate({
      options: {
        name: validBucketName,
        planId: validPlanId,
        ownerGroupId: validOwnerGroupId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when roster id is used with owner group name', async () => {
    const actual = await command.validate({
      options: {
        name: validBucketName,
        rosterId: validRosterId,
        ownerGroupName: validOwnerGroupName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when roster id is used with owner group id', async () => {
    const actual = await command.validate({
      options: {
        name: validBucketName,
        rosterId: validRosterId,
        ownerGroupId: validOwnerGroupId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with id', async () => {
    const actual = await command.validate({
      options: {
        id: validBucketId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct input with name', async () => {
    const actual = await command.validate({
      options: {
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when no groups found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {

      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'&$select=id`) {
        return { "value": [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    }), new CommandError(`The specified group '${validOwnerGroupName}' does not exist.`));
  });

  it('fails validation when multiple groups found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {

      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'&$select=id`) {
        return multipleGroupResponse;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    }), new CommandError("Multiple groups with name 'Group name' found. Found: 00000000-0000-0000-0000-000000000000."));
  });

  it('fails validation when no buckets found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return { "value": [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: validBucketName,
        planId: validPlanId
      }
    }), new CommandError(`The specified bucket '${validBucketName}' does not exist.`));
  });

  it('fails validation when multiple buckets found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return multipleBucketByNameResponse;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: validBucketName,
        planId: validPlanId
      }
    }), new CommandError("Multiple buckets with name 'Bucket name' found. Found: vncYUXCRBke28qMLB-d4xJcACtNz."));
  });

  it('handles selecting single result when multiple groups with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'&$select=id`) {
        return singleGroupResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans?$select=id,title`) {
        return singlePlanResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return multipleBucketByNameResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}`) {
        return singleBucketByIdResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves(singleBucketByNameResponse.value[0]);

    await assert.doesNotReject(command.action(logger, {
      options: {
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    }));
  });

  it('correctly gets bucket by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}`) {
        return singleBucketByIdResponse;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        id: validBucketId
      }
    }));
  });

  it('correctly gets bucket by name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {

      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'&$select=id`) {
        return singleGroupResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans?$select=id,title`) {
        return singlePlanResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return singleBucketByNameResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}`) {
        return singleBucketByIdResponse;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    }));
  });

  it('correctly gets bucket by plan title and owner group ID', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans?$select=id,title`) {
        return singlePlanResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return singleBucketByNameResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}`) {
        return singleBucketByIdResponse;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupId: validOwnerGroupId
      }
    }));
  });

  it('correctly gets bucket by roster id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/plans?$select=id`) {
        return planResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return singleBucketByNameResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}`) {
        return singleBucketByIdResponse;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, {
      options: {
        name: validBucketName,
        rosterId: validRosterId
      }
    } as any);
    assert(loggerLogSpy.calledWith(singleBucketByIdResponse));
  });
});
