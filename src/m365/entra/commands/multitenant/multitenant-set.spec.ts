import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
import { cli } from '../../../../cli/cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import commands from '../../commands';
const command: Command = require('./multitenant-set');
import { telemetry } from '../../../../telemetry';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import request from '../../../../request';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';

describe(commands.MULTITENANT_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
    (command as any).pollingInterval = 0;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MULTITENANT_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when only displayName is specified', async () => {
    const actual = await command.validate({ options: { displayName: 'Contoso organization' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when only description is specified', async () => {
    const actual = await command.validate({ options: { description: 'Contoso and partners' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the displayName and description are specified', async () => {
    const actual = await command.validate({ options: { displayName: 'Contoso organization', description: 'Contoso and partners' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when no option is specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('updates a displayName of a multitenant organization', async () => {
    const patchRequestStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: 'Contoso organization' } });
    assert(patchRequestStub.called);
  });

  it('updates a description of a multitenant organization', async () => {
    const patchRequestStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { description: 'Contoso and partners', verbose: true } });
    assert(patchRequestStub.called);
  });

  it('updates displayName and description of a multitenant organization', async () => {
    const patchRequestStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: 'Contoso organization', description: 'Contoso and partners' } });
    assert(patchRequestStub.called);
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'patch').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'Invalid request'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: {} }), new CommandError('Invalid request'));
  });
});