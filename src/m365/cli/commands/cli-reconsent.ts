import { cli } from '../../../cli/cli';
import { Logger } from '../../../cli/Logger';
import { settingsNames } from '../../../settingsNames';
import { browserUtil } from '../../../utils/browserUtil';
import AnonymousCommand from '../../base/AnonymousCommand';
import commands from '../commands';

class CliReconsentCommand extends AnonymousCommand {
  public get name(): string {
    return commands.RECONSENT;
  }

  public get description(): string {
    return 'Returns URL to open in the browser to re-consent CLI for Microsoft 365 Microsoft Entra permissions';
  }

  public async commandAction(logger: Logger): Promise<void> {
    const url = `https://login.microsoftonline.com/${cli.getTenant()}/oauth2/authorize?client_id=${cli.getClientId()}&response_type=code&prompt=admin_consent`;

    if (cli.getSettingWithDefaultValue<boolean>(settingsNames.autoOpenLinksInBrowser, false) === false) {
      await logger.log(`To re-consent your Microsoft Entra application, navigate in your web browser to ${url}.`);
      return;
    }

    await logger.log(`Opening the following page in your browser: ${url}`);

    try {
      await browserUtil.open(url);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new CliReconsentCommand();