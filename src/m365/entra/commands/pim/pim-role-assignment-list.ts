import { UnifiedRoleAssignmentScheduleInstance } from '@microsoft/microsoft-graph-types';
import GlobalOptions from '../../../../GlobalOptions';
import { Logger } from '../../../../cli/Logger';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { validation } from '../../../../utils/validation';
import { entraUser } from '../../../../utils/entraUser';
import { entraGroup } from '../../../../utils/entraGroup';
import { odata } from '../../../../utils/odata';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  userId?: string;
  userName?: string;
  groupId?: string;
  groupName?: string;
  startDateTime?: string;
  includePrincipalDetails?: boolean;
}

class EntraPimRoleAssignmentListCommand extends GraphCommand {
  public get name(): string {
    return commands.PIM_ROLE_ASSIGNMENT_LIST;
  }

  public get description(): string {
    return 'Retrieves a list of Entra role assignments for a user or group';
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
        userId: typeof args.options.userId !== 'undefined',
        userName: typeof args.options.userName !== 'undefined',
        groupId: typeof args.options.groupId !== 'undefined',
        groupName: typeof args.options.groupName !== 'undefined',
        startDateTime: typeof args.options.startDateTime !== 'undefined',
        includePrincipalDetails: !!args.options.includePrincipalDetails
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: "--userId [userId]"
      },
      {
        option: "--userName [userName]"
      },
      {
        option: "--groupId [groupId]"
      },
      {
        option: "--groupName [groupName]"
      },
      {
        option: "-s, --startDateTime [startDateTime]"
      },
      {
        option: "--includePrincipalDetails [includePrincipalDetails]"
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.userId && !validation.isValidGuid(args.options.userId)) {
          return `${args.options.userId} is not a valid GUID`;
        }

        if (args.options.groupId && !validation.isValidGuid(args.options.groupId)) {
          return `${args.options.groupId} is not a valid GUID`;
        }

        if (args.options.startDateTime && !validation.isValidISODateTime(args.options.startDateTime)) {
          return `${args.options.startDateTime} is not a valid ISO 8601 date time string`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({
      options: ['userId', 'userName', 'groupId', 'groupName'],
      runsWhen: (args) => args.options.userId || args.options.userName || args.options.groupId || args.options.groupName
    });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const queryParameters: string[] = [];
    const filters: string[] = [];
    const expands: string[] = [];

    try {
      const principalId = await this.getPrincipalId(logger, args.options);

      if (principalId) {
        filters.push(`principalId eq '${principalId}'`);
      }

      if (args.options.startDateTime) {
        filters.push(`startDateTime ge ${args.options.startDateTime}`);
      }

      if (filters.length > 0) {
        queryParameters.push(`$filter=${filters.join(' and ')}`);
      }

      expands.push('roleDefinition($select=displayName)');

      if (args.options.includePrincipalDetails) {
        expands.push('principal');
      }

      queryParameters.push(`$expand=${expands.join(',')}`);

      const queryString = `?${queryParameters.join('&')}`;

      const url = `${this.resource}/v1.0/roleManagement/directory/roleAssignmentScheduleInstances${queryString}`;

      const results = await odata.getAllItems<UnifiedRoleAssignmentScheduleInstance>(url);

      await logger.log(results);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getPrincipalId(logger: Logger, options: Options): Promise<string | undefined> {
    let principalId = options.userId;

    if (options.userName) {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving user by its name '${options.userName}'`);
      }

      principalId = await entraUser.getUserIdByUpn(options.userName);
    }
    else if (options.groupId) {
      principalId = options.groupId;
    }
    else if (options.groupName) {
      if (this.verbose) {
        await logger.logToStderr(`Retrieving group by its name '${options.groupName}'`);
      }

      principalId = await entraGroup.getGroupIdByDisplayName(options.groupName);
    }

    return principalId;
  }
}

module.exports = new EntraPimRoleAssignmentListCommand();