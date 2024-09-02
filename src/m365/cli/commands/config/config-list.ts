import { cli } from "../../../../cli/cli";
import { Logger } from "../../../../cli/Logger";
import AnonymousCommand from "../../../base/AnonymousCommand";
import commands from "../../commands";

class CliConfigListCommand extends AnonymousCommand {
  public get name(): string {
    return commands.CONFIG_LIST;
  }

  public get description(): string {
    return 'List all self set CLI for Microsoft 365 configurations';
  }

  public async commandAction(logger: Logger): Promise<void> {
    await logger.log(cli.getConfig().all);
  }
}

module.exports = new CliConfigListCommand();