import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
import { cli } from '../../../../cli/cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { telemetry } from '../../../../telemetry';
import { accessToken } from '../../../../utils/accessToken';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./info-get');

describe(commands.INFO_GET, () => {
  const domainName = 'contoso.com';
  const tenantId = 'e65b162c-6f87-4eb1-a24e-1b37d3504663';
  const tenantInfoResponse = {
    tenantId: tenantId,
    federationBrandName: null,
    displayName: "Contoso",
    defaultDomainName: domainName
  };

  let log: any[];
  let loggerLogSpy: sinon.SinonSpy;
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    if (!auth.connection.accessTokens[auth.defaultResource]) {
      auth.connection.accessTokens[auth.defaultResource] = {
        expiresOn: '123',
        accessToken: 'abc'
      };
    }
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.INFO_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the tenantId is not a valid guid', async () => {
    const actual = await command.validate({ options: { tenantId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the tenantId is a valid GUID', async () => {
    const actual = await command.validate({ options: { tenantId: tenantId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if both domainName and tenantId are specified', async () => {
    const actual = await command.validate({ options: { domainName: domainName, tenantId: tenantId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('gets tenant information for the currently signed in user if no domain name or tenantId is passed', async () => {
    sinon.stub(accessToken, 'getUserNameFromAccessToken').callsFake(() => {
      return 'admin@contoso.com';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/tenantRelationships/findTenantInformationByDomainName(domainName='contoso.com')`) {
        return tenantInfoResponse;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledOnceWithExactly(tenantInfoResponse));
    sinonUtil.restore(accessToken.getUserNameFromAccessToken);
  });

  it('gets tenant information with correct domain name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/tenantRelationships/findTenantInformationByDomainName(domainName='contoso.com')`) {
        return tenantInfoResponse;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { verbose: true, domainName: domainName } });
    assert(loggerLogSpy.calledOnceWithExactly(tenantInfoResponse));
  });

  it('gets tenant information with correct tenant id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/tenantRelationships/findTenantInformationByTenantId(tenantId='e65b162c-6f87-4eb1-a24e-1b37d3504663')`) {
        return tenantInfoResponse;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, { options: { verbose: true, tenantId: tenantId } });
    assert(loggerLogSpy.calledOnceWithExactly(tenantInfoResponse));
  });

  it('handles error when trying to retrieve information for a non-existant tenant by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/tenantRelationships/findTenantInformationByTenantId(tenantId='e65b162c-6f87-4eb1-a24e-1b37d3504663')`) {
        throw {
          "error": {
            "code": "Directory_ObjectNotFound",
            "message": "Unable to read the company information from the directory.",
            "innerError": {
              "date": "2023-09-14T14:07:47",
              "request-id": "3b91132c-5c79-454b-8dd4-06964e788a24",
              "client-request-id": "2147e6c6-8036-cc2f-f4d0-eec89dbc48d7"
            }
          }
        };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, { options: { tenantId: tenantId } } as any), new CommandError("Unable to read the company information from the directory."));
  });

  it('handles error when trying to retrieve information for a non-existant tenant by domain name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/tenantRelationships/findTenantInformationByDomainName(domainName='xyz.com')`) {
        throw {
          "error": {
            "code": "Directory_ObjectNotFound",
            "message": "Unable to read the company information from the directory.",
            "innerError": {
              "date": "2023-09-14T14:07:47",
              "request-id": "3b91132c-5c79-454b-8dd4-06964e788a24",
              "client-request-id": "2147e6c6-8036-cc2f-f4d0-eec89dbc48d7"
            }
          }
        };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, { options: { domainName: 'xyz.com' } } as any), new CommandError("Unable to read the company information from the directory."));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));
    await assert.rejects(command.action(logger, { options: { domainName: 'xyz.com' } } as any), new CommandError('An error has occurred'));
  });
});