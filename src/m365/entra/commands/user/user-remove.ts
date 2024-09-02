import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import commands from '../../commands';
import request, { CliRequestOptions } from '../../../../request';
import { validation } from '../../../../utils/validation';
import { cli } from '../../../../cli/cli';
import GraphCommand from '../../../base/GraphCommand';
import aadCommands from '../../aadCommands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  userName?: string;
  force?: boolean;
}

class EntraUserRemoveCommand extends GraphCommand {

  public get name(): string {
    return commands.USER_REMOVE;
  }

  public get description(): string {
    return 'Removes a specific user';
  }

  public alias(): string[] | undefined {
    return [aadCommands.USER_REMOVE];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        id: typeof args.options.id !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--id [id]'
      },
      {
        option: '--userName [userName]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      { options: ['id', 'userName'] }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.id && !validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        if (args.options.userName && !validation.isValidUserPrincipalName(args.options.userName)) {
          return `${args.options.userName} is not a valid user principal name (UPN)`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: any): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.USER_REMOVE, commands.USER_REMOVE);

    if (this.verbose) {
      await logger.logToStderr(`Removing user '${args.options.id || args.options.userName}'...`);
    }

    if (args.options.force) {
      await this.deleteUser(args);
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove user '${args.options.id || args.options.userName}'?` });

      if (result) {
        await this.deleteUser(args);
      }
    }
  }

  private async deleteUser(args: CommandArgs): Promise<void> {
    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/users/${args.options.id || args.options.userName}`,
        headers: {
          accept: 'application/json;odata.metadata=none'
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

module.exports = new EntraUserRemoveCommand();