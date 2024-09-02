import auth from '../../../Auth';
import { Logger } from '../../../cli/Logger';
import GlobalOptions from '../../../GlobalOptions';
import { urlUtil } from '../../../utils/urlUtil';
import { validation } from '../../../utils/validation';
import SpoCommand from '../../base/SpoCommand';
import commands from '../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
}

class SpoSetCommand extends SpoCommand {
  public get name(): string {
    return commands.SET;
  }

  public get description(): string {
    return 'Sets the URL of the root SharePoint site collection for use in SPO commands';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --url <url>'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.url)
    );
  }

  #initTypes(): void {
    this.types.string.push('url');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    auth.connection.spoUrl = urlUtil.removeTrailingSlashes(args.options.url);

    try {
      await auth.storeConnectionInfo();
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSetCommand();