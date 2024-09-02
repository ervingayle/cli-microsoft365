import { Logger } from '../../../../cli/Logger';
import { odata } from '../../../../utils/odata';
import GraphCommand from '../../../base/GraphCommand';
import aadCommands from '../../aadCommands';
import commands from '../../commands';

class EntraLicenseListCommand extends GraphCommand {
  public get name(): string {
    return commands.LICENSE_LIST;
  }

  public get description(): string {
    return 'Lists commercial subscriptions that an organization has acquired';
  }

  public alias(): string[] | undefined {
    return [aadCommands.LICENSE_LIST];
  }

  public defaultProperties(): string[] | undefined {
    return ['id', 'skuId', 'skuPartNumber'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    await this.showDeprecationWarning(logger, aadCommands.LICENSE_LIST, commands.LICENSE_LIST);

    if (this.verbose) {
      await logger.logToStderr(`Retrieving the commercial subscriptions that an organization has acquired`);
    }

    try {
      const items = await odata.getAllItems<any>(`${this.resource}/v1.0/subscribedSkus`);
      await logger.log(items);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new EntraLicenseListCommand();