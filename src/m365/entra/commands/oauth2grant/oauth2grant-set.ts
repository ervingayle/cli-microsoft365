import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import GraphCommand from '../../../base/GraphCommand';
import aadCommands from '../../aadCommands';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  grantId: string;
  scope: string;
}

class EntraOAuth2GrantSetCommand extends GraphCommand {
  public get name(): string {
    return commands.OAUTH2GRANT_SET;
  }

  public get description(): string {
    return 'Update OAuth2 permissions for the service principal';
  }

  public alias(): string[] | undefined {
    return [aadCommands.OAUTH2GRANT_SET];
  }

  constructor() {
    super();

    this.#initOptions();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --grantId <grantId>'
      },
      {
        option: '-s, --scope <scope>'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.OAUTH2GRANT_SET, commands.OAUTH2GRANT_SET);

    if (this.verbose) {
      await logger.logToStderr(`Updating OAuth2 permissions...`);
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/v1.0/oauth2PermissionGrants/${formatting.encodeQueryParameter(args.options.grantId)}`,
        headers: {
          'content-type': 'application/json'
        },
        responseType: 'json',
        data: {
          "scope": args.options.scope
        }
      };

      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new EntraOAuth2GrantSetCommand();