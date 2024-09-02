import { cli } from '../../../../cli/cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { teams } from '../../../../utils/teams';
import { validation } from '../../../../utils/validation';
import DelegatedGraphCommand from '../../../base/DelegatedGraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId?: string;
  teamName?: string;
  channelId?: string;
  channelName?: string;
  id: string;
  force?: boolean;
}

class TeamsMessageRemoveCommand extends DelegatedGraphCommand {
  public get name(): string {
    return commands.MESSAGE_REMOVE;
  }

  public get description(): string {
    return 'Removes a message from a channel in a Microsoft Teams team';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        teamId: typeof args.options.teamId !== 'undefined',
        teamName: typeof args.options.teamName !== 'undefined',
        channelId: typeof args.options.channelId !== 'undefined',
        channelName: typeof args.options.channelName !== 'undefined',
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--teamId [teamId]'
      },
      {
        option: '--teamName [teamName]'
      },
      {
        option: '--channelId [channelId]'
      },
      {
        option: '--channelName [channelName]'
      },
      {
        option: '-i, --id <id>'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
          return `'${args.options.teamId}' is not a valid GUID for 'teamId'.`;
        }

        if (args.options.channelId && !validation.isValidTeamsChannelId(args.options.channelId)) {
          return `'${args.options.channelId}' is not a valid ID for 'channelId'.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push(
      {
        options: ['teamId', 'teamName']
      },
      {
        options: ['channelId', 'channelName']
      }
    );
  }

  #initTypes(): void {
    this.types.string.push('teamId', 'teamName', 'channelId', 'channelName', 'id');
    this.types.boolean.push('force');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removeTeamMessage = async (): Promise<void> => {
      try {
        if (this.verbose) {
          await logger.logToStderr(`Removing message '${args.options.id}' from channel '${args.options.channelId || args.options.channelName}' in team '${args.options.teamId || args.options.teamName}'.`);
        }

        const teamId: string = args.options.teamId || await teams.getTeamIdByDisplayName(args.options.teamName!);
        const channelId: string = args.options.channelId || await teams.getChannelIdByDisplayName(teamId, args.options.channelName!);

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/teams/${teamId}/channels/${formatting.encodeQueryParameter(channelId)}/messages/${args.options.id}/softDelete`,
          headers: {
            accept: 'application/json;odata.metadata=none'
          },
          responseType: 'json'
        };

        await request.post(requestOptions);
      }
      catch (err: any) {
        if (err.error?.error?.code === 'NotFound') {
          this.handleError('The message was not found in the Teams channel.');
        }
        else {
          this.handleRejectedODataJsonPromise(err);
        }
      }
    };

    if (args.options.force) {
      await removeTeamMessage();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove this message?` });

      if (result) {
        await removeTeamMessage();
      }
    }
  }
}

module.exports = new TeamsMessageRemoveCommand();