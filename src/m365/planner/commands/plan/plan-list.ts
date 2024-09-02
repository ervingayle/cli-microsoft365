import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { entraGroup } from '../../../../utils/entraGroup';
import { planner } from '../../../../utils/planner';
import { validation } from '../../../../utils/validation';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  ownerGroupId?: string;
  ownerGroupName?: string;
  rosterId?: string;
}

class PlannerPlanListCommand extends GraphCommand {
  public get name(): string {
    return commands.PLAN_LIST;
  }

  public get description(): string {
    return 'Returns a list of plans associated with a specified group or roster';
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
        ownerGroupId: typeof args.options.ownerGroupId !== 'undefined',
        ownerGroupName: typeof args.options.ownerGroupName !== 'undefined',
        rosterId: typeof args.options.rosterId !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: "--ownerGroupId [ownerGroupId]"
      },
      {
        option: "--ownerGroupName [ownerGroupName]"
      },
      {
        option: "--rosterId [rosterId]"
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.ownerGroupId && !validation.isValidGuid(args.options.ownerGroupId)) {
          return `${args.options.ownerGroupId} is not a valid GUID`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['ownerGroupId', 'ownerGroupName', 'rosterId'] });
  }

  #initTypes(): void {
    this.types.string.push('ownerGroupId', 'ownerGroupName', 'rosterId');
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'title', 'createdDateTime', 'owner'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let plannerPlans = [];
      if (args.options.ownerGroupId || args.options.ownerGroupName) {
        const groupId = await this.getGroupId(args);
        plannerPlans = await planner.getPlansByGroupId(groupId);
      }
      else {
        const plan = await planner.getPlanByRosterId(args.options.rosterId!);
        plannerPlans.push(plan);
      }

      await logger.log(plannerPlans);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getGroupId(args: CommandArgs): Promise<string> {
    if (args.options.ownerGroupId) {
      return args.options.ownerGroupId;
    }

    return entraGroup.getGroupIdByDisplayName(args.options.ownerGroupName!);
  }
}

module.exports = new PlannerPlanListCommand();