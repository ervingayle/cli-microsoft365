import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import GraphCommand from '../../../base/GraphCommand';
import request from '../../../../request';
import { accessToken } from '../../../../utils/accessToken';
import { formatting } from '../../../../utils/formatting';
import { validation } from '../../../../utils/validation';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  domainName?: string;
  tenantId?: string;
}

class TenantInfoGetCommand extends GraphCommand {
  public get name(): string {
    return commands.INFO_GET;
  }

  public get description(): string {
    return 'Gets information about any tenant';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        domainName: typeof args.options.domainName !== 'undefined',
        tenantId: typeof args.options.tenantId !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-d, --domainName [domainName]'
      },
      {
        option: '-i, --tenantId [tenantId]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.tenantId && !validation.isValidGuid(args.options.tenantId)) {
          return `${args.options.tenantId} is not a valid GUID`;
        }

        if (args.options.tenantId && args.options.domainName) {
          return `Specify either domainName or tenantId but not both`;
        }

        return true;
      }
    );
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    let domainName: string | undefined = args.options.domainName;
    const tenantId: string | undefined = args.options.tenantId;

    if (!domainName && !tenantId) {
      const userName: string = accessToken.getUserNameFromAccessToken(auth.connection.accessTokens[auth.defaultResource].accessToken);
      domainName = userName.split('@')[1];
    }

    let requestUrl = `${this.resource}/v1.0/tenantRelationships/`;

    if (tenantId) {
      requestUrl += `findTenantInformationByTenantId(tenantId='${formatting.encodeQueryParameter(tenantId)}')`;
    }
    else {
      requestUrl += `findTenantInformationByDomainName(domainName='${formatting.encodeQueryParameter(domainName!)}')`;
    }

    const requestOptions: any = {
      url: requestUrl,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    try {
      const res: any = await request.get(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new TenantInfoGetCommand();