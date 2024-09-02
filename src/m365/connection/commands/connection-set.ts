import { Logger } from '../../../cli/Logger';
import auth from '../../../Auth';
import commands from '../commands';
import Command, { CommandError } from '../../../Command';
import GlobalOptions from '../../../GlobalOptions';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  newName: string;
}

class ConnectionSetCommand extends Command {
  public get name(): string {
    return commands.SET;
  }

  public get description(): string {
    return 'Rename the specified connection';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '--newName <newName>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.name === args.options.newName) {
          return `Choose a name different from the current one`;
        }

        return true;
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('name', 'newName');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const connection = await auth.getConnection(args.options.name);

    if (this.verbose) {
      await logger.logToStderr(`Updating connection '${connection.identityName}', appId: ${connection.appId}, tenantId: ${connection.identityTenantId}...`);
    }

    await auth.updateConnection(connection, args.options.newName);
  }

  public async action(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      await auth.restoreAuth();
    }
    catch (error: any) {
      throw new CommandError(error);
    }

    await this.initAction(args, logger);
    await this.commandAction(logger, args);
  }
}

module.exports = new ConnectionSetCommand();