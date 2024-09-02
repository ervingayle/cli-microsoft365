import Command, { CommandArgs, CommandError } from "../../Command";
import auth from "../../Auth";
import { Logger } from "../../cli/Logger";
import { accessToken } from "../../utils/accessToken";

export default abstract class VivaEngageCommand extends Command {
  protected get resource(): string {
    return 'https://www.yammer.com/api';
  }

  protected async initAction(args: CommandArgs, logger: Logger): Promise<void> {
    await super.initAction(args, logger);

    if (!auth.connection.active) {
      // we fail no login in the base command command class
      return;
    }

    accessToken.assertDelegatedAccessToken();
  }

  protected handleRejectedODataJsonPromise(response: any): void {
    if (response.statusCode === 404) {
      throw new CommandError("Not found (404)");
    }
    else if (response.error && response.error.base) {
      throw new CommandError(response.error.base);
    }
    else {
      throw new CommandError(response);
    }
  }
}