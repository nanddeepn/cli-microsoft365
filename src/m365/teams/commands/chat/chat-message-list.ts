import { ItemBody, Message } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import { validation } from '../../../../utils/validation.js';
import GraphCommand from '../../../base/GraphCommand.js';
import commands from '../../commands.js';
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';

interface ExtendedMessage extends Message {
  shortBody?: string;
}

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  chatId: z.string()
    .refine(id => validation.isValidTeamsChatId(id), {
      error: e => `'${e.input}' is not a valid value for option chatId.`
    })
    .alias('i'),
  createdEndDateTime: z.string()
    .refine(time => validation.isValidISODateTime(time), {
      error: e => `'${e.input}' is not a valid ISO date-time string for option createdEndDateTime.`
    })
    .optional(),
  endDateTime: z.string()
    .refine(time => validation.isValidISODateTime(time), {
      error: e => `'${e.input}' is not a valid ISO date-time string for option endDateTime.`
    })
    .optional(),
  modifiedStartDateTime: z.string()
    .refine(time => validation.isValidISODateTime(time), {
      error: e => `'${e.input}' is not a valid ISO date-time string for option modifiedStartDateTime.`
    })
    .optional(),
  modifiedEndDateTime: z.string()
    .refine(time => validation.isValidISODateTime(time), {
      error: e => `'${e.input}' is not a valid ISO date-time string for option modifiedEndDateTime.`
    })
    .optional()
});

declare type Options = z.infer<typeof options>;
interface CommandArgs {
  options: Options;
}

class TeamsChatMessageListCommand extends GraphCommand {
  public get name(): string {
    return commands.CHAT_MESSAGE_LIST;
  }

  public get description(): string {
    return 'Lists all messages from a chat';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'createdDateTime', 'shortBody'];
  }

  public get schema(): z.ZodType {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
    return schema
      .refine(options => !(options.endDateTime && options.createdEndDateTime), {
        error: 'Specify either endDateTime or createdEndDateTime, but not both.'
      })
      .refine(options => !(options.createdEndDateTime && (options.modifiedStartDateTime || options.modifiedEndDateTime)), {
        error: 'You cannot combine createdEndDateTime with modifiedStartDateTime or modifiedEndDateTime. These filters operate on different properties.'
      })
      .refine(options => !(options.endDateTime && (options.modifiedStartDateTime || options.modifiedEndDateTime)), {
        error: 'You cannot combine endDateTime with modifiedStartDateTime or modifiedEndDateTime. These filters operate on different properties.'
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.endDateTime) {
      await this.warn(logger, `Option 'endDateTime' is deprecated. Please use 'createdEndDateTime' instead.`);

      args.options.createdEndDateTime = args.options.endDateTime;
    }

    try {
      let apiUrl = `${this.resource}/v1.0/chats/${args.options.chatId}/messages`;

      if (args.options.createdEndDateTime) {
        apiUrl += `?$filter=createdDateTime lt ${args.options.createdEndDateTime}&$orderby=createdDateTime desc`;
      }
      else if (args.options.modifiedStartDateTime || args.options.modifiedEndDateTime) {
        const filters: string[] = [];

        if (args.options.modifiedStartDateTime) {
          filters.push(`lastModifiedDateTime gt ${args.options.modifiedStartDateTime}`);
        }

        if (args.options.modifiedEndDateTime) {
          filters.push(`lastModifiedDateTime lt ${args.options.modifiedEndDateTime}`);
        }

        apiUrl += `?$filter=${filters.join(' and ')}&$orderby=lastModifiedDateTime desc`;
      }

      const items = await odata.getAllItems<ExtendedMessage>(apiUrl);
      if (args.options.output && args.options.output !== 'json') {
        items.forEach(i => {
          // hoist the content to body for readability
          i.body = (i.body as ItemBody).content as any;

          let shortBody: string | undefined;
          const bodyToProcess = i.body as string;

          if (bodyToProcess) {
            let maxLength = 50;
            let addedDots = '...';
            if (bodyToProcess.length < maxLength) {
              maxLength = bodyToProcess.length;
              addedDots = '';
            }

            shortBody = bodyToProcess.replace(/\n/g, ' ').substring(0, maxLength) + addedDots;
          }

          i.shortBody = shortBody;
        });
      }

      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new TeamsChatMessageListCommand();