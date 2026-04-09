import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
import GraphCommand from '../../../base/GraphCommand.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { validation } from '../../../../utils/validation.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import auth from '../../../../Auth.js';
import { formatting } from '../../../../utils/formatting.js';
import { calendarGroup } from '../../../../utils/calendarGroup.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  id: z.string().optional(),
  name: z.string().optional(),
  userId: z.string().refine(id => validation.isValidGuid(id), {
    error: e => `'${e.input}' is not a valid GUID.`
  }).optional(),
  userName: z.string().refine(name => validation.isValidUserPrincipalName(name), {
    error: e => `'${e.input}' is not a valid UPN.`
  }).optional(),
  newName: z.string()
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class OutlookCalendarGroupSetCommand extends GraphCommand {
  public get name(): string {
    return commands.CALENDARGROUP_SET;
  }

  public get description(): string {
    return 'Updates a calendar group for a user';
  }

  public get schema(): z.ZodType | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => !(options.userId && options.userName), {
        error: 'Specify either userId or userName, but not both.'
      })
      .refine(options => !(!options.id && !options.name), {
        error: 'Specify either id or name.'
      })
      .refine(options => !(options.id && options.name), {
        error: 'Specify either id or name, but not both.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      const token = auth.connection.accessTokens[auth.defaultResource].accessToken;
      const isAppOnlyAccessToken = accessToken.isAppOnlyAccessToken(token);

      let userUrl: string;
      let graphUserId: string;

      if (isAppOnlyAccessToken) {
        if (!args.options.userId && !args.options.userName) {
          throw 'When running with application permissions either userId or userName is required.';
        }

        graphUserId = (args.options.userId ?? args.options.userName)!;
        userUrl = `${this.resource}/v1.0/users('${formatting.encodeQueryParameter(graphUserId)}')`;

        if (this.verbose) {
          await logger.logToStderr(`Updating calendar group using application permissions for user '${graphUserId}'...`);
        }
      }
      else if (args.options.userId || args.options.userName) {
        const currentUserId = accessToken.getUserIdFromAccessToken(token);
        const currentUserName = accessToken.getUserNameFromAccessToken(token);
        const isOtherUser = (args.options.userId && args.options.userId !== currentUserId) ||
          (args.options.userName && args.options.userName.toLowerCase() !== currentUserName?.toLowerCase());

        if (isOtherUser) {
          const scopes = accessToken.getScopesFromAccessToken(token);
          const hasSharedScope = scopes.some(s => s === 'Calendars.ReadWrite.Shared');

          if (!hasSharedScope) {
            throw `To update calendar groups of other users, the Entra ID application used for authentication must have the Calendars.ReadWrite.Shared delegated permission assigned.`;
          }
        }

        graphUserId = (args.options.userId ?? args.options.userName)!;
        userUrl = `${this.resource}/v1.0/users('${formatting.encodeQueryParameter(graphUserId)}')`;

        if (this.verbose) {
          await logger.logToStderr(`Updating calendar group using delegated permissions for user '${graphUserId}'...`);
        }
      }
      else {
        graphUserId = accessToken.getUserIdFromAccessToken(token);
        userUrl = `${this.resource}/v1.0/me`;

        if (this.verbose) {
          await logger.logToStderr('Updating calendar group for the signed-in user...');
        }
      }

      let calendarGroupId: string;
      if (args.options.id) {
        calendarGroupId = args.options.id;
      }
      else {
        const calendarGroupResult = await calendarGroup.getUserCalendarGroupByName(graphUserId, args.options.name!);
        calendarGroupId = calendarGroupResult.id!;
      }

      if (this.verbose) {
        await logger.logToStderr(`Updating calendar group '${calendarGroupId}'...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `${userUrl}/calendarGroups/${calendarGroupId}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        responseType: 'json',
        data: {
          name: args.options.newName
        }
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new OutlookCalendarGroupSetCommand();
