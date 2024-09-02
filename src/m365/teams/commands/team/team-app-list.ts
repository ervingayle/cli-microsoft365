import { cli } from '../../../../cli/cli';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import * as teamGetCommand from './team-get';
import { Options as TeamsTeamGetOptions } from './team-get';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId?: string;
  teamName?: string;
}

class TeamsTeamAppListCommand extends GraphCommand {
  public get name(): string {
    return commands.TEAM_APP_LIST;
  }

  public get description(): string {
    return 'List apps installed in the specified team';
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'displayName', 'distributionMethod'];
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        teamId: typeof args.options.teamId !== 'undefined',
        teamName: typeof args.options.teamName !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --teamId [teamId]'
      },
      {
        option: '-n, --teamName [teamName]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.teamId && !validation.isValidGuid(args.options.teamId)) {
          return `${args.options.teamId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['teamId', 'teamName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving installed apps for team '${args.options.teamId || args.options.teamName}'`);
      }

      const teamId: string = await this.getTeamId(args);
      const res = await odata.getAllItems<any>(`${this.resource}/v1.0/teams/${teamId}/installedApps?$expand=teamsApp,teamsAppDefinition`);

      if (!cli.shouldTrimOutput(args.options.output)) {
        await logger.log(res);
      }
      else {
        //converted to text friendly output
        await logger.log(res.map(i => {
          return {
            id: i.id,
            displayName: i.teamsApp.displayName,
            distributionMethod: i.teamsApp.distributionMethod
          };
        }));
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTeamId(args: CommandArgs): Promise<string> {
    if (args.options.teamId) {
      return args.options.teamId;
    }

    const teamGetOptions: TeamsTeamGetOptions = {
      name: args.options.teamName,
      debug: this.debug,
      verbose: this.verbose
    };

    const commandOutput = await cli.executeCommandWithOutput(teamGetCommand as Command, { options: { ...teamGetOptions, _: [] } });
    const team = JSON.parse(commandOutput.stdout);
    return team.id;
  }
}

module.exports = new TeamsTeamAppListCommand();