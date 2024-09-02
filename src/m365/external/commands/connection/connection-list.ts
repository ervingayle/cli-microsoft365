import { Logger } from '../../../../cli/Logger';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';

class ExternalConnectionListCommand extends GraphCommand {
  public get name(): string {
    return commands.CONNECTION_LIST;
  }

  public get description(): string {
    return 'Lists external connections defined in the Microsoft Search';
  }

  public alias(): string[] | undefined {
    return [commands.EXTERNALCONNECTION_LIST];
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'name', 'state'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    try {
      const connections = await odata.getAllItems(`${this.resource}/v1.0/external/connections`);
      await logger.log(connections);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new ExternalConnectionListCommand();