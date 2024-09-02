import { cli } from '../../../../cli/cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folderUrl: string;
  force?: boolean;
}

class SpoFolderRoleInheritanceResetCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_ROLEINHERITANCE_RESET;
  }

  public get description(): string {
    return 'Restores the role inheritance of a folder';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '--folderUrl <folderUrl>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  #initTypes(): void {
    this.types.string.push('webUrl', 'folderUrl');
    this.types.boolean.push('force');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const serverRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.folderUrl);
    const roleFolderUrl: string = urlUtil.getWebRelativePath(args.options.webUrl, args.options.folderUrl);
    let requestUrl: string = `${args.options.webUrl}/_api/web/`;

    const resetFolderRoleInheritance = async (): Promise<void> => {
      try {
        if (roleFolderUrl.split('/').length === 2) {
          requestUrl += `GetList('${formatting.encodeQueryParameter(serverRelativeUrl)}')`;
        }
        else {
          requestUrl += `GetFolderByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')/ListItemAllFields`;
        }
        const requestOptions: CliRequestOptions = {
          url: `${requestUrl}/resetroleinheritance`,
          headers: {
            accept: 'application/json;odata=nometadata'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await resetFolderRoleInheritance();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to reset the role inheritance of folder ${args.options.folderUrl} located in site ${args.options.webUrl}?` });

      if (result) {
        await resetFolderRoleInheritance();
      }
    }
  }
}

module.exports = new SpoFolderRoleInheritanceResetCommand();