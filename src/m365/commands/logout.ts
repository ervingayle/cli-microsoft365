import auth from '../../Auth';
import { Logger } from '../../cli/Logger';
import Command, { CommandArgs, CommandError } from '../../Command';
import commands from './commands';

class LogoutCommand extends Command {
  public get name(): string {
    return commands.LOGOUT;
  }

  public get description(): string {
    return 'Log out from Microsoft 365';
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Logging out from Microsoft 365...');
    }

    const deactivate: () => void = (): void => auth.connection.deactivate();

    try {
      await auth.clearConnectionInfo();
    }
    catch (error: any) {
      if (this.debug) {
        await logger.logToStderr(new CommandError(error));
      }
    }
    finally {
      deactivate();
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
    await this.commandAction(logger);
  }
}

module.exports = new LogoutCommand();