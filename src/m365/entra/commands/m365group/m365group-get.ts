import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { entraGroup } from '../../../../utils/entraGroup';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import aadCommands from '../../aadCommands';
import commands from '../../commands';
import { GroupExtended } from './GroupExtended';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  includeSiteUrl: boolean;
}

class EntraM365GroupGetCommand extends GraphCommand {
  public get name(): string {
    return commands.M365GROUP_GET;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft 365 Group or Microsoft Teams team';
  }

  public alias(): string[] | undefined {
    return [aadCommands.M365GROUP_GET];
  }

  constructor() {
    super();

    this.#initOptions();
    this.#initValidators();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '--includeSiteUrl'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (!validation.isValidGuid(args.options.id)) {
          return `${args.options.id} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.M365GROUP_GET, commands.M365GROUP_GET);

    let group: GroupExtended;

    try {
      const isUnifiedGroup = await entraGroup.isUnifiedGroup(args.options.id);

      if (!isUnifiedGroup) {
        throw Error(`Specified group with id '${args.options.id}' is not a Microsoft 365 group.`);
      }

      group = await entraGroup.getGroupById(args.options.id);

      if (args.options.includeSiteUrl) {
        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/groups/${group.id}/drive?$select=webUrl`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        const res = await request.get<{ webUrl: string }>(requestOptions);
        group.siteUrl = res.webUrl ? res.webUrl.substr(0, res.webUrl.lastIndexOf('/')) : '';
      }

      await logger.log(group);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new EntraM365GroupGetCommand();
