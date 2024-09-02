import { Logger } from '../../../../cli/Logger';
import request, { CliRequestOptions } from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

class PlannerRosterAddCommand extends GraphCommand {
  public get name(): string {
    return commands.ROSTER_ADD;
  }

  public get description(): string {
    return 'Creates a new Microsoft Planner Roster';
  }

  public async commandAction(logger: Logger): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr('Creating a new Microsoft Planner Roster');
    }

    try {
      const requestOptions: CliRequestOptions = {
        url: `${this.resource}/beta/planner/rosters`,
        headers: {
          accept: 'application/json;odata.metadata=none'
        },
        data: {},
        responseType: 'json'
      };

      const response = await request.post(requestOptions);
      await logger.log(response);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

}

module.exports = new PlannerRosterAddCommand();