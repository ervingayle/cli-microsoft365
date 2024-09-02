import { cli } from '../../../../cli/cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import VivaEngageCommand from "../../../base/VivaEngageCommand";
import commands from '../../commands';
import yammerCommands from './yammerCommands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  groupId: number;
  id?: number;
  force?: boolean;
}

class VivaEngageGroupUserRemoveCommand extends VivaEngageCommand {
  public get name(): string {
    return commands.ENGAGE_GROUP_USER_REMOVE;
  }

  public get description(): string {
    return 'Removes a user from a Viva Engage group';
  }

  public alias(): string[] | undefined {
    return [yammerCommands.GROUP_USER_REMOVE];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        userId: args.options.id !== undefined,
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--groupId <groupId>'
      },
      {
        option: '--id [id]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.groupId && typeof args.options.groupId !== 'number') {
          return `${args.options.groupId} is not a number`;
        }

        if (args.options.id && typeof args.options.id !== 'number') {
          return `${args.options.id} is not a number`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, this.alias()![0], this.name);

    if (args.options.force) {
      await this.executeRemoveAction(args.options);
    }
    else {
      let messagePrompt: string = `Are you sure you want to leave group ${args.options.groupId}?`;
      if (args.options.id) {
        messagePrompt = `Are you sure you want to remove the user ${args.options.id} from the group ${args.options.groupId}?`;
      }

      const result = await cli.promptForConfirmation({ message: messagePrompt });

      if (result) {
        await this.executeRemoveAction(args.options);
      }
    }
  }

  private async executeRemoveAction(options: GlobalOptions): Promise<void> {
    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1/group_memberships.json`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        // eslint-disable-next-line camelcase
        group_id: options.groupId,
        // eslint-disable-next-line camelcase
        user_id: options.id
      }
    };

    try {
      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new VivaEngageGroupUserRemoveCommand();