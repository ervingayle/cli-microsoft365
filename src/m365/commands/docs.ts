import { cli } from '../../cli/cli';
import { Logger } from '../../cli/Logger';
import { settingsNames } from '../../settingsNames';
import { app } from '../../utils/app';
import { browserUtil } from '../../utils/browserUtil';
import AnonymousCommand from '../base/AnonymousCommand';
import commands from './commands';

class DocsCommand extends AnonymousCommand {
  public get name(): string {
    return commands.DOCS;
  }

  public get description(): string {
    return 'Returns the CLI for Microsoft 365 docs webpage URL';
  }

  public async commandAction(logger: Logger): Promise<void> {
    await logger.log(app.packageJson().homepage);

    if (cli.getSettingWithDefaultValue<boolean>(settingsNames.autoOpenLinksInBrowser, false) === false) {
      return;
    }

    await browserUtil.open(app.packageJson().homepage);
  }
}

module.exports = new DocsCommand();