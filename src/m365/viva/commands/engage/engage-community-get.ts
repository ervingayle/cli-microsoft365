import GlobalOptions from '../../../../GlobalOptions';
import { Logger } from '../../../../cli/Logger';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import request, { CliRequestOptions } from '../../../../request';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class VivaEngageCommunityGetCommand extends GraphCommand {
  public get name(): string {
    return commands.ENGAGE_COMMUNITY_GET;
  }

  public get description(): string {
    return 'Gets information of a Viva Engage community';
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initTypes();
  }

  #initOptions(): void {
    this.options.unshift(
      { option: '-i, --id <id>' }
    );
  }

  #initTypes(): void {
    this.types.string.push('id');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Getting the information of Viva Engage community with id '${args.options.id}'...`);
    }

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/beta/employeeExperience/communities/${args.options.id}`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<any>(requestOptions);

      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new VivaEngageCommunityGetCommand();