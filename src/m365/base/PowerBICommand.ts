import Command, { CommandArgs } from '../../Command';
import auth from '../../Auth';
import { accessToken } from '../../utils/accessToken';
import { Logger } from '../../cli/Logger';

export default abstract class PowerBICommand extends Command {
  protected get resource(): string {
    return 'https://api.powerbi.com';
  }

  protected async initAction(args: CommandArgs, logger: Logger): Promise<void> {
    await super.initAction(args, logger);

    if (!auth.connection.active) {
      // we fail no login in the base command command class
      return;
    }

    accessToken.assertDelegatedAccessToken();
  }

}
