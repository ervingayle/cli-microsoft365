import auth, { CloudType } from '../../Auth';
import { Logger } from '../../cli/Logger';
import Command, { CommandArgs, CommandError } from '../../Command';
import { accessToken } from '../../utils/accessToken';


export default abstract class PowerPlatformCommand extends Command {
  protected get resource(): string {
    return 'https://api.bap.microsoft.com';
  }

  protected async initAction(args: CommandArgs, logger: Logger): Promise<void> {
    await super.initAction(args, logger);

    if (!auth.connection.active) {
      // we fail no login in the base command command class
      return;
    }

    if (auth.connection.cloudType !== CloudType.Public) {
      throw new CommandError(`Power Platform commands only support the public cloud at the moment. We'll add support for other clouds in the future. Sorry for the inconvenience.`);
    }

    accessToken.assertDelegatedAccessToken();
  }
}
