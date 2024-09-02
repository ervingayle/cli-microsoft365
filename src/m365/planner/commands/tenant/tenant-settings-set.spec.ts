import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
import { cli } from '../../../../cli/cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { telemetry } from '../../../../telemetry';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./tenant-settings-set');

describe(commands.TENANT_SETTINGS_SET, () => {
  const successResponse = {
    id: '1',
    isPlannerAllowed: true,
    allowCalendarSharing: true,
    allowTenantMoveWithDataLoss: false,
    allowTenantMoveWithDataMigration: false,
    allowRosterCreation: true,
    allowPlannerMobilePushNotifications: true
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_SETTINGS_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation no options are specified', async () => {
    const actual = await command.validate({
      options: {}
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });


  it('passes validation when valid options specified', async () => {
    const actual = await command.validate({
      options: {
        isPlannerAllowed: 'true',
        allowCalendarSharing: 'false',
        allowPlannerMobilePushNotifications: 'false'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('successfully updates tenant planner settings', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://tasks.office.com/taskAPI/tenantAdminSettings/Settings') {
        return successResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        isPlannerAllowed: 'true'
      }
    });
    assert(loggerLogSpy.calledWith(successResponse));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://tasks.office.com/taskAPI/tenantAdminSettings/Settings') {
        throw 'An error has occurred';
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        isPlannerAllowed: 'true'
      }
    }), new CommandError('An error has occurred'));
  });
});
