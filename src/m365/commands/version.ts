import { Logger } from '../../cli/Logger';
import { app } from '../../utils/app';
import AnonymousCommand from '../base/AnonymousCommand';
import commands from './commands';

class VersionCommand extends AnonymousCommand {
  public get name(): string {
    return commands.VERSION;
  }

  public get description(): string {
    return 'Shows CLI for Microsoft 365 version';
  }

  public async commandAction(logger: Logger): Promise<void> {
    await logger.log(`v${app.packageJson().version}`);
  }
}

module.exports = new VersionCommand();