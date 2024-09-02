import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import aadCommands from '../../aadCommands';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  password: string;
}

class EntraUserPasswordValidateCommand extends GraphCommand {
  public get name(): string {
    return commands.USER_PASSWORD_VALIDATE;
  }

  public get description(): string {
    return "Check a user's password against the organization's password validation policy";
  }

  public alias(): string[] | undefined {
    return [aadCommands.USER_PASSWORD_VALIDATE];
  }

  constructor() {
    super();

    this.#initOptions();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-p, --password <password>'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.USER_PASSWORD_VALIDATE, commands.USER_PASSWORD_VALIDATE);

    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/users/validatePassword`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        data: {
          password: args.options.password
        },
        responseType: 'json'
      };

      const res = await request.post(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new EntraUserPasswordValidateCommand();