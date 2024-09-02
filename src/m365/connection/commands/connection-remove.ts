import { Logger } from '../../../cli/Logger';
import auth from '../../../Auth';
import commands from '../commands';
import Command, { CommandError } from '../../../Command';
import GlobalOptions from '../../../GlobalOptions';
import { cli } from '../../../cli/cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  force?: boolean;
}

class ConnectionRemoveCommand extends Command {
  public get name(): string {
    return commands.REMOVE;
  }

  public get description(): string {
    return 'Remove the specified connection';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initTypes();
  }


  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-n, --name <name>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('name');
    this.types.boolean.push('force');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const deleteConnection = async (): Promise<void> => {
      const connection = await auth.getConnection(args.options.name);

      if (this.verbose) {
        await logger.logToStderr(`Removing connection '${connection.identityName}', appId: ${connection.appId}, tenantId: ${connection.identityTenantId}...`);
      }

      await auth.removeConnectionInfo(connection, logger, this.debug);
    };

    if (args.options.force) {
      await deleteConnection();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the connection?` });

      if (result) {
        await deleteConnection();
      }
    }
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

module.exports = new ConnectionRemoveCommand();