import { Conversation } from '@microsoft/microsoft-graph-types';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { entraGroup } from '../../../../utils/entraGroup';
import aadCommands from '../../aadCommands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId: string;
}

class EntraM365GroupConversationListCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_CONVERSATION_LIST;
  }

  public get description(): string {
    return 'Lists conversations for the specified Microsoft 365 group';
  }

  public alias(): string[] | undefined {
    return [aadCommands.M365GROUP_CONVERSATION_LIST];
  }

  public defaultProperties(): string[] | undefined {
    return ['topic', 'lastDeliveredDateTime', 'id'];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --groupId <groupId>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.groupId as string)) {
          return `${args.options.groupId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.M365GROUP_CONVERSATION_LIST, commands.M365GROUP_CONVERSATION_LIST);

    try {
      const isUnifiedGroup = await entraGroup.isUnifiedGroup(args.options.groupId);

      if (!isUnifiedGroup) {
        throw Error(`Specified group with id '${args.options.groupId}' is not a Microsoft 365 group.`);
      }

      const conversations = await odata.getAllItems<Conversation>(`${this.resource}/v1.0/groups/${args.options.groupId}/conversations`);
      await logger.log(conversations);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new EntraM365GroupConversationListCommand();