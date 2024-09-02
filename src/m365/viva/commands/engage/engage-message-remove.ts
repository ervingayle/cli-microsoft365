import { cli } from '../../../../cli/cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import VivaEngageCommand from '../../../base/VivaEngageCommand';
import commands from '../../commands';
import yammerCommands from './yammerCommands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: number;
  force?: boolean;
}

class VivaEngageMessageRemoveCommand extends VivaEngageCommand {
  public get name(): string {
    return commands.ENGAGE_MESSAGE_REMOVE;
  }

  public get description(): string {
    return 'Removes a Viva Engage message';
  }

  public alias(): string[] | undefined {
    return [yammerCommands.MESSAGE_REMOVE];
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
        force: (!(!args.options.force)).toString()
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id <id>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (typeof args.options.id !== 'number') {
          return `${args.options.id} is not a number`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, this.alias()![0], this.name);

    if (args.options.force) {
      await this.removeMessage(args.options);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the Viva Engage message ${args.options.id}?` });

      if (result) {
        await this.removeMessage(args.options);
      }
    }
  }

  private async removeMessage(options: Options): Promise<void> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1/messages/${options.id}.json`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'content-type': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      await request.delete(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new VivaEngageMessageRemoveCommand();
