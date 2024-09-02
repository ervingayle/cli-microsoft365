import { Logger } from '../../../../cli/Logger';
import request, { CliRequestOptions } from '../../../../request';
import GraphCommand from '../../../base/GraphCommand';
import commands from '../../commands';
import { MultitenantOrganization } from './MultitenantOrganization';

class EntraMultitenantGetCommand extends GraphCommand {
  public get name(): string {
    return commands.MULTITENANT_GET;
  }

  public get description(): string {
    return 'Gets properties of the multitenant organization';
  }

  public async commandAction(logger: Logger): Promise<void> {

    const requestOptions: CliRequestOptions = {
      url: `${this.resource}/v1.0/tenantRelationships/multiTenantOrganization`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      if (this.verbose) {
        await logger.logToStderr('Retrieving multitenant organization...');
      }
      const multitenantOrg = await request.get<MultitenantOrganization>(requestOptions);

      await logger.log(multitenantOrg);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new EntraMultitenantGetCommand();