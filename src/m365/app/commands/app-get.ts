import { cli } from '../../../cli/cli';
import { Logger } from '../../../cli/Logger';
import Command from '../../../Command';
import * as entraAppGetCommand from '../../entra/commands/app/app-get';
import { Options as EntraAppGetCommandOptions } from '../../entra/commands/app/app-get';
import AppCommand, { AppCommandArgs } from '../../base/AppCommand';
import commands from '../commands';

class AppGetCommand extends AppCommand {
  public get name(): string {
    return commands.GET;
  }

  public get description(): string {
    return 'Retrieves information about the current Microsoft Entra app';
  }

  public async commandAction(logger: Logger, args: AppCommandArgs): Promise<void> {
    const options: EntraAppGetCommandOptions = {
      appId: this.appId,
      output: 'json',
      debug: args.options.debug,
      verbose: args.options.verbose
    };

    try {
      const appGetOutput = await cli.executeCommandWithOutput(entraAppGetCommand as Command, { options: { ...options, _: [] } });
      if (this.verbose) {
        await logger.logToStderr(appGetOutput.stderr);
      }

      await logger.log(JSON.parse(appGetOutput.stdout));
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new AppGetCommand();