import { PlannerPlan } from '@microsoft/microsoft-graph-types';
import { cli } from '../../../../cli/cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { entraGroup } from '../../../../utils/entraGroup';
import { planner } from '../../../../utils/planner';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id?: string;
  title?: string;
  ownerGroupId?: string;
  ownerGroupName?: string;
  force?: boolean;
}

class PlannerPlanRemoveCommand extends GraphCommand {
  public get name(): string {
    return commands.PLAN_REMOVE;
  }

  public get description(): string {
    return 'Removes the Microsoft Planner plan';
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
        id: typeof args.options.id !== 'undefined',
        title: typeof args.options.title !== 'undefined',
        ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
        ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
        force: !!args.options.force
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id [id]'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '--ownerGroupId [ownerGroupId]'
      },
      {
        option: '--ownerGroupName [ownerGroupName]'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.title) {
          if (!args.options.ownerGroupId && !args.options.ownerGroupName) {
            return 'Specify either ownerGroupId or ownerGroupName';
          }

          if (args.options.ownerGroupId && args.options.ownerGroupName) {
            return 'Specify either ownerGroupId or ownerGroupName but not both';
          }

          if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
            return `${args.options.ownerGroupId} is not a valid GUID`;
          }
        }
        else if (args.options.ownerGroupId || args.options.ownerGroupName) {
          return 'Don\'t specify ownerGroupId or ownerGroupName when using id';
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['id', 'title'] });
  }

  #initTypes(): void {
    this.types.string.push('id', 'title', 'ownerGroupId', 'ownerGroupName');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const removePlan = async (): Promise<void> => {
      try {
        const plan = await this.getPlan(args);

        if (this.verbose) {
          await logger.logToStderr(`Removing plan '${plan.title}' ...`);
        }

        const requestOptions: CliRequestOptions = {
          url: `${this.resource}/v1.0/planner/plans/${plan.id}`,
          headers: {
            accept: 'application/json;odata.metadata=none',
            'if-match': (plan as any)['@odata.etag']
          },
          responseType: 'json'
        };

        await request.delete(requestOptions);
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removePlan();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove the plan ${args.options.id || args.options.title}?` });

      if (result) {
        await removePlan();
      }
    }
  }

  private async getPlan(args: CommandArgs): Promise<PlannerPlan> {
    const { id, title } = args.options;

    if (id) {
      return planner.getPlanById(id, 'minimal');
    }

    const groupId = await this.getGroupId(args);
    return planner.getPlanByTitle(title!, groupId, 'minimal');
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    const { ownerGroupId, ownerGroupName } = args.options;

    if (ownerGroupId) {
      return ownerGroupId;
    }

    return entraGroup.getGroupIdByDisplayName(ownerGroupName!);
  }
}

module.exports = new PlannerPlanRemoveCommand();