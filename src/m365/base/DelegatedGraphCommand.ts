import auth from '../../Auth';
import { CommandArgs } from '../../Command';
import { Logger } from '../../cli/Logger';
import { accessToken } from '../../utils/accessToken';
import GraphCommand from './GraphCommand';

/**
 * This command class is for delegated-only Graph commands.  
 */
export default abstract class DelegatedGraphCommand extends GraphCommand {
  protected async initAction(args: CommandArgs, logger: Logger): Promise<void> {
    await super.initAction(args, logger);

    if (!auth.connection.active) {
      // we fail no login in the base command command class
      return;
    }

    accessToken.assertDelegatedAccessToken();
  }
}