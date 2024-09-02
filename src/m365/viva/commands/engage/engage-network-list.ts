import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import VivaEngageCommand from '../../../base/VivaEngageCommand';
import commands from '../../commands';
import yammerCommands from './yammerCommands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  includeSuspended: boolean;
}

class VivaEngageNetworkListCommand extends VivaEngageCommand {
  public get name(): string {
    return commands.ENGAGE_NETWORK_LIST;
  }

  public get description(): string {
    return 'Returns a list of networks to which the current user has access';
  }

  public alias(): string[] | undefined {
    return [yammerCommands.NETWORK_LIST];
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'email', 'community', 'permalink', 'web_url'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        includeSuspended: args.options.includeSuspended
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--includeSuspended'
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, this.alias()![0], this.name);

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1/networks/current.json`,
      headers: {
        accept: 'application/json;odata.metadata=none',
        'content-type': 'application/json;odata=nometadata'
      },
      responseType: 'json',
      data: {
        includeSuspended: args.options.includeSuspended !== undefined && args.options.includeSuspended !== false
      }
    };

    try {
      const res: any = await request.get(requestOptions);

      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new VivaEngageNetworkListCommand();