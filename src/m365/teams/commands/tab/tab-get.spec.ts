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
import command from './tab-get.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.TAB_GET, () => {
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TAB_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both teamId and teamName options are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        channelId: '19:00000000000000000000000000000000@thread.skype',
        id: '00000000-0000-0000-0000-000000000000'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both teamId and teamName options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        teamId: '26b48cd6-3da7-493d-8010-1b246ef552d6',
        teamName: 'Team Name',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        id: '00000000-0000-0000-0000-000000000000'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both channelId and channelName options are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        teamId: '26b48cd6-3da7-493d-8010-1b246ef552d6',
        id: '00000000-0000-0000-0000-000000000000'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both channelId and channelName options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        teamId: '26b48cd6-3da7-493d-8010-1b246ef552d6',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        channelName: 'Channel Name',
        id: '00000000-0000-0000-0000-000000000000'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and name options are not passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        teamId: '26b48cd6-3da7-493d-8010-1b246ef552d6',
        channelId: '19:00000000000000000000000000000000@thread.skype'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and name options are passed', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        teamId: '26b48cd6-3da7-493d-8010-1b246ef552d6',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        id: '00000000-0000-0000-0000-000000000000',
        name: 'Tab Name'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        id: '00000000-0000-0000-0000-000000000000'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not provided.', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        channelId: '19:00000000000000000000000000000000@thread.skype',
        id: '00000000-0000-0000-0000-000000000000'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the channelId is not provided.', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validates for a incorrect channelId missing leading 19:.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '00000000000000000000000000000000@thread.skype',
        name: 'Tab'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validates for a incorrect channelId missing trailing @thread.skype.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread',
        name: 'Tab'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        id: '00000000-0000'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id is not provided.', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        name: 'Tab Name'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly handles teams tabs request failure due to wrong channel id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/29%3A00000000000000000000000000000000%40thread.skype/tabs/00000000-0000-0000-0000-000000000000?$expand=teamsApp') {
        throw {
          "error": {
            "code": "Invalid request",
            "message": "Channel id is not in a valid format: 29:00000000000000000000000000000000@thread.skype",
            "innerError": {
              "request-id": "75c4e0f1-035e-47e3-917b-0c8823a02a96",
              "date": "2020-07-19T11:08:32"
            }
          }
        };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '29:00000000000000000000000000000000@thread.skype',
        id: '00000000-0000-0000-0000-000000000000'
      }
    } as any), new CommandError('Channel id is not in a valid format: 29:00000000000000000000000000000000@thread.skype'));
  });

  it('should get a Microsoft Teams Tab by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19%3A00000000000000000000000000000000%40thread.skype/tabs/00000000-0000-0000-0000-000000000000?$expand=teamsApp`) {
        return {
          "id": "00000000-0000-0000-0000-000000000000",
          "displayName": "TeamsTab",
          "webUrl": "https://teams.microsoft.com/l/entity/00000000-0000-0000-0000-000000000000/_djb2_msteams_prefix_00000000-0000-0000-0000-000000000000?label=TeamsTab&context=%7b%0d%0a++%22canvasUrl%22%3a+%22https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fTeamsLogon.aspx%3fSPFX%3dtrue%26dest%3d%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fteamshostedapp.aspx%253Flist%3d7d7f911a-bf19-46a0-86d9-187c3f32cce2%2526id%3d2%2526webPartInstanceId%3d1c8e5fda-7fd7-416f-9930-b3e90f009ea5%22%2c%0d%0a++%22channelId%22%3a+%2219%3000000000000000000000000000000008%40thread.skype%22%2c%0d%0a++%22subEntityId%22%3a+null%0d%0a%7d&groupId=00000000-0000-0000-0000-000000000000&tenantId=de348bc7-1aeb-4406-8cb3-97db021cadb4",
          "configuration": {
            "entityId": "sharepointtab_00000000-0000-0000-0000-000000000000",
            "contentUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5",
            "removeUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5%26removeTab",
            "websiteUrl": null,
            "dateAdded": "2020-07-18T19:27:22.03Z"
          },
          "teamsApp": { "id": "2a527703-1f6f-4559-a332-d8a7d288cd88", "externalId": null, "displayName": "Website", "distributionMethod": "store" }
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        id: '00000000-0000-0000-0000-000000000000'
      }
    });

    assert(loggerLogSpy.calledWith({
      "id": "00000000-0000-0000-0000-000000000000",
      "displayName": "TeamsTab",
      "webUrl": "https://teams.microsoft.com/l/entity/00000000-0000-0000-0000-000000000000/_djb2_msteams_prefix_00000000-0000-0000-0000-000000000000?label=TeamsTab&context=%7b%0d%0a++%22canvasUrl%22%3a+%22https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fTeamsLogon.aspx%3fSPFX%3dtrue%26dest%3d%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fteamshostedapp.aspx%253Flist%3d7d7f911a-bf19-46a0-86d9-187c3f32cce2%2526id%3d2%2526webPartInstanceId%3d1c8e5fda-7fd7-416f-9930-b3e90f009ea5%22%2c%0d%0a++%22channelId%22%3a+%2219%3000000000000000000000000000000008%40thread.skype%22%2c%0d%0a++%22subEntityId%22%3a+null%0d%0a%7d&groupId=00000000-0000-0000-0000-000000000000&tenantId=de348bc7-1aeb-4406-8cb3-97db021cadb4",
      "configuration": {
        "entityId": "sharepointtab_00000000-0000-0000-0000-000000000000",
        "contentUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5",
        "removeUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5%26removeTab",
        "websiteUrl": null,
        "dateAdded": "2020-07-18T19:27:22.03Z"
      },
      "teamsApp": { "id": "2a527703-1f6f-4559-a332-d8a7d288cd88", "externalId": null, "displayName": "Website", "distributionMethod": "store" }
    }));
  });

  it('fails when team name does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq 'Team%20Name'`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 1,
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000000",
              "createdDateTime": null,
              "displayName": "Team Name",
              "description": "Team Description",
              "internalId": null,
              "classification": null,
              "specialization": null,
              "visibility": null,
              "webUrl": null,
              "isArchived": false,
              "isMembershipLimitedToOwners": null,
              "memberSettings": null,
              "guestSettings": null,
              "messagingSettings": null,
              "funSettings": null,
              "discoverySettings": null,
              "resourceProvisioningOptions": []
            }
          ]
        };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        teamName: 'Team Name',
        channelName: 'Channel Name',
        name: 'Tab Name'
      }
    } as any), new CommandError('The specified team does not exist in the Microsoft Teams'));
  });

  it('should get a Microsoft Teams Tab by Team name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq 'Team%20Name'`) {
        return {
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000000",
              "createdDateTime": null,
              "displayName": "Team Name",
              "description": "Team Description",
              "internalId": null,
              "classification": null,
              "specialization": null,
              "visibility": null,
              "webUrl": null,
              "isArchived": false,
              "isMembershipLimitedToOwners": null,
              "memberSettings": null,
              "guestSettings": null,
              "messagingSettings": null,
              "funSettings": null,
              "discoverySettings": null,
              "resourceProvisioningOptions": ["Team"]
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels?$filter=displayName eq 'Channel%20Name'`) {
        return {
          "value": [
            {
              "id": "19:00000000-0000-0000-0000-000000000000",
              "displayName": "Channel Name",
              "description": "Channel description",
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a00000000000000000000000000000000%40thread.tacv2/Channel+Name?groupId=00000000-0000-0000-0000-000000000000&tenantId=00000000-0000-0000-0000-000000000000",
              "membershipType": "standard"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19%3A00000000-0000-0000-0000-000000000000/tabs?$filter=displayName eq 'Tab%20Name'`) {
        return {
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000000",
              "displayName": "TeamsTab",
              "webUrl": "https://teams.microsoft.com/l/entity/00000000-0000-0000-0000-000000000000/_djb2_msteams_prefix_00000000-0000-0000-0000-000000000000?label=TeamsTab&context=%7b%0d%0a++%22canvasUrl%22%3a+%22https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fTeamsLogon.aspx%3fSPFX%3dtrue%26dest%3d%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fteamshostedapp.aspx%253Flist%3d7d7f911a-bf19-46a0-86d9-187c3f32cce2%2526id%3d2%2526webPartInstanceId%3d1c8e5fda-7fd7-416f-9930-b3e90f009ea5%22%2c%0d%0a++%22channelId%22%3a+%2219%3000000000000000000000000000000008%40thread.skype%22%2c%0d%0a++%22subEntityId%22%3a+null%0d%0a%7d&groupId=00000000-0000-0000-0000-000000000000&tenantId=de348bc7-1aeb-4406-8cb3-97db021cadb4",
              "configuration": {
                "entityId": "sharepointtab_00000000-0000-0000-0000-000000000000",
                "contentUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5",
                "removeUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5%26removeTab",
                "websiteUrl": null,
                "dateAdded": "2020-07-18T19:27:22.03Z"
              }
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19%3A00000000-0000-0000-0000-000000000000/tabs/00000000-0000-0000-0000-000000000000?$expand=teamsApp`) {
        return {
          "id": "00000000-0000-0000-0000-000000000000",
          "displayName": "TeamsTab",
          "webUrl": "https://teams.microsoft.com/l/entity/00000000-0000-0000-0000-000000000000/_djb2_msteams_prefix_00000000-0000-0000-0000-000000000000?label=TeamsTab&context=%7b%0d%0a++%22canvasUrl%22%3a+%22https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fTeamsLogon.aspx%3fSPFX%3dtrue%26dest%3d%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fteamshostedapp.aspx%253Flist%3d7d7f911a-bf19-46a0-86d9-187c3f32cce2%2526id%3d2%2526webPartInstanceId%3d1c8e5fda-7fd7-416f-9930-b3e90f009ea5%22%2c%0d%0a++%22channelId%22%3a+%2219%3000000000000000000000000000000008%40thread.skype%22%2c%0d%0a++%22subEntityId%22%3a+null%0d%0a%7d&groupId=00000000-0000-0000-0000-000000000000&tenantId=de348bc7-1aeb-4406-8cb3-97db021cadb4",
          "configuration": {
            "entityId": "sharepointtab_00000000-0000-0000-0000-000000000000",
            "contentUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5",
            "removeUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5%26removeTab",
            "websiteUrl": null,
            "dateAdded": "2020-07-18T19:27:22.03Z"
          },
          "teamsApp": { "id": "2a527703-1f6f-4559-a332-d8a7d288cd88", "externalId": null, "displayName": "Website", "distributionMethod": "store" }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        teamName: 'Team Name',
        channelName: 'Channel Name',
        name: 'Tab Name'
      }
    });

    assert(loggerLogSpy.calledWith({
      "id": "00000000-0000-0000-0000-000000000000",
      "displayName": "TeamsTab",
      "webUrl": "https://teams.microsoft.com/l/entity/00000000-0000-0000-0000-000000000000/_djb2_msteams_prefix_00000000-0000-0000-0000-000000000000?label=TeamsTab&context=%7b%0d%0a++%22canvasUrl%22%3a+%22https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fTeamsLogon.aspx%3fSPFX%3dtrue%26dest%3d%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fteamshostedapp.aspx%253Flist%3d7d7f911a-bf19-46a0-86d9-187c3f32cce2%2526id%3d2%2526webPartInstanceId%3d1c8e5fda-7fd7-416f-9930-b3e90f009ea5%22%2c%0d%0a++%22channelId%22%3a+%2219%3000000000000000000000000000000008%40thread.skype%22%2c%0d%0a++%22subEntityId%22%3a+null%0d%0a%7d&groupId=00000000-0000-0000-0000-000000000000&tenantId=de348bc7-1aeb-4406-8cb3-97db021cadb4",
      "configuration": {
        "entityId": "sharepointtab_00000000-0000-0000-0000-000000000000",
        "contentUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5",
        "removeUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5%26removeTab",
        "websiteUrl": null,
        "dateAdded": "2020-07-18T19:27:22.03Z"
      },
      "teamsApp": { "id": "2a527703-1f6f-4559-a332-d8a7d288cd88", "externalId": null, "displayName": "Website", "distributionMethod": "store" }
    }));
  });

  it('should get a Microsoft Teams Tab by Channel name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels?$filter=displayName eq 'Channel%20Name'`) {
        return {
          "value": [
            {
              "id": "19:00000000-0000-0000-0000-000000000000",
              "displayName": "Channel Name",
              "description": "Channel description",
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a00000000000000000000000000000000%40thread.tacv2/Channel+Name?groupId=00000000-0000-0000-0000-000000000000&tenantId=00000000-0000-0000-0000-000000000000",
              "membershipType": "standard"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19%3A00000000-0000-0000-0000-000000000000/tabs?$filter=displayName eq 'Tab%20Name'`) {
        return {
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000000",
              "displayName": "TeamsTab",
              "webUrl": "https://teams.microsoft.com/l/entity/00000000-0000-0000-0000-000000000000/_djb2_msteams_prefix_00000000-0000-0000-0000-000000000000?label=TeamsTab&context=%7b%0d%0a++%22canvasUrl%22%3a+%22https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fTeamsLogon.aspx%3fSPFX%3dtrue%26dest%3d%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fteamshostedapp.aspx%253Flist%3d7d7f911a-bf19-46a0-86d9-187c3f32cce2%2526id%3d2%2526webPartInstanceId%3d1c8e5fda-7fd7-416f-9930-b3e90f009ea5%22%2c%0d%0a++%22channelId%22%3a+%2219%3000000000000000000000000000000008%40thread.skype%22%2c%0d%0a++%22subEntityId%22%3a+null%0d%0a%7d&groupId=00000000-0000-0000-0000-000000000000&tenantId=de348bc7-1aeb-4406-8cb3-97db021cadb4",
              "configuration": {
                "entityId": "sharepointtab_00000000-0000-0000-0000-000000000000",
                "contentUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5",
                "removeUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5%26removeTab",
                "websiteUrl": null,
                "dateAdded": "2020-07-18T19:27:22.03Z"
              }
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19%3A00000000-0000-0000-0000-000000000000/tabs/00000000-0000-0000-0000-000000000000?$expand=teamsApp`) {
        return {
          "id": "00000000-0000-0000-0000-000000000000",
          "displayName": "TeamsTab",
          "webUrl": "https://teams.microsoft.com/l/entity/00000000-0000-0000-0000-000000000000/_djb2_msteams_prefix_00000000-0000-0000-0000-000000000000?label=TeamsTab&context=%7b%0d%0a++%22canvasUrl%22%3a+%22https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fTeamsLogon.aspx%3fSPFX%3dtrue%26dest%3d%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fteamshostedapp.aspx%253Flist%3d7d7f911a-bf19-46a0-86d9-187c3f32cce2%2526id%3d2%2526webPartInstanceId%3d1c8e5fda-7fd7-416f-9930-b3e90f009ea5%22%2c%0d%0a++%22channelId%22%3a+%2219%3000000000000000000000000000000008%40thread.skype%22%2c%0d%0a++%22subEntityId%22%3a+null%0d%0a%7d&groupId=00000000-0000-0000-0000-000000000000&tenantId=de348bc7-1aeb-4406-8cb3-97db021cadb4",
          "configuration": {
            "entityId": "sharepointtab_00000000-0000-0000-0000-000000000000",
            "contentUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5",
            "removeUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5%26removeTab",
            "websiteUrl": null,
            "dateAdded": "2020-07-18T19:27:22.03Z"
          },
          "teamsApp": { "id": "2a527703-1f6f-4559-a332-d8a7d288cd88", "externalId": null, "displayName": "Website", "distributionMethod": "store" }
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        teamId: '00000000-0000-0000-0000-000000000000',
        channelName: 'Channel Name',
        name: 'Tab Name'
      }
    });

    assert(loggerLogSpy.calledWith({
      "id": "00000000-0000-0000-0000-000000000000",
      "displayName": "TeamsTab",
      "webUrl": "https://teams.microsoft.com/l/entity/00000000-0000-0000-0000-000000000000/_djb2_msteams_prefix_00000000-0000-0000-0000-000000000000?label=TeamsTab&context=%7b%0d%0a++%22canvasUrl%22%3a+%22https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fTeamsLogon.aspx%3fSPFX%3dtrue%26dest%3d%2fsites%2fPrototypeTeam%2f_layouts%2f15%2fteamshostedapp.aspx%253Flist%3d7d7f911a-bf19-46a0-86d9-187c3f32cce2%2526id%3d2%2526webPartInstanceId%3d1c8e5fda-7fd7-416f-9930-b3e90f009ea5%22%2c%0d%0a++%22channelId%22%3a+%2219%3000000000000000000000000000000008%40thread.skype%22%2c%0d%0a++%22subEntityId%22%3a+null%0d%0a%7d&groupId=00000000-0000-0000-0000-000000000000&tenantId=de348bc7-1aeb-4406-8cb3-97db021cadb4",
      "configuration": {
        "entityId": "sharepointtab_00000000-0000-0000-0000-000000000000",
        "contentUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5",
        "removeUrl": "https://contoso.sharepoint.com/sites/PrototypeTeam/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/sites/PrototypeTeam/_layouts/15/teamshostedapp.aspx%3Flist=7d7f911a-bf19-46a0-86d9-187c3f32cce2%26id=2%26webPartInstanceId=1c8e5fda-7fd7-416f-9930-b3e90f009ea5%26removeTab",
        "websiteUrl": null,
        "dateAdded": "2020-07-18T19:27:22.03Z"
      },
      "teamsApp": { "id": "2a527703-1f6f-4559-a332-d8a7d288cd88", "externalId": null, "displayName": "Website", "distributionMethod": "store" }
    }));
  });

  it('fails to get channel when channel does not exists', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels?$filter=displayName eq 'Channel%20Name'`) {
        return { value: [] };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        teamId: '00000000-0000-0000-0000-000000000000',
        channelName: 'Channel Name',
        name: 'Tab Name'
      }
    } as any), new CommandError('The specified channel does not exist in the Microsoft Teams team'));
  });

  it('fails to get tab when tab does not exists', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19%3A00000000000000000000000000000000%40thread.skype/tabs?$filter=displayName eq 'Tab%20Name'`) {
        return { value: [] };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:00000000000000000000000000000000@thread.skype',
        name: 'Tab Name'
      }
    } as any), new CommandError('The specified tab does not exist in the Microsoft Teams team channel'));
  });
});
