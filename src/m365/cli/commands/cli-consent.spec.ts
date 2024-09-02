import * as assert from 'assert';
import * as sinon from 'sinon';
import { cli } from '../../../cli/cli';
import { CommandInfo } from '../../../cli/CommandInfo';
import { Logger } from '../../../cli/Logger';
import { telemetry } from '../../../telemetry';
import { pid } from '../../../utils/pid';
import { session } from '../../../utils/session';
import commands from '../commands';
const command: Command = require('./cli-consent');
import { sinonUtil } from '../../../utils/sinonUtil';
import Command from '../../../Command';

describe(commands.CONSENT, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: any;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
  });

  afterEach(() => {
    sinonUtil.restore([
      cli.getTenant,
      cli.getClientId,
      (command as any).warn
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONSENT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('shows consent URL for VivaEngage permissions for a custom single-tenant app', async () => {
    sinon.stub(cli, 'getTenant').returns('fb5cb38f-ecdb-4c6a-a93b-b8cfd56b4a89');
    sinon.stub(cli, 'getClientId').returns('2587b55d-a41e-436d-bb1d-6223eb185dd4');
    await command.action(logger, { options: { service: 'VivaEngage' } });
    assert(loggerLogSpy.calledWith(`To consent permissions for executing VivaEngage commands, navigate in your web browser to https://login.microsoftonline.com/fb5cb38f-ecdb-4c6a-a93b-b8cfd56b4a89/oauth2/v2.0/authorize?client_id=2587b55d-a41e-436d-bb1d-6223eb185dd4&response_type=code&scope=https%3A%2F%2Fapi.yammer.com%2Fuser_impersonation`));
  });

  it('shows warning for Yammer permissions', async () => {
    const warnSpy = sinon.spy(command as any, 'warn');
    await command.action(logger, { options: { service: 'yammer' } });
    assert(warnSpy.called);
  });

  it('supports specifying service', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--service') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if specified service is invalid ', async () => {
    const actual = await command.validate({ options: { service: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if service is set to VivaEngage ', async () => {
    const actual = await command.validate({ options: { service: 'VivaEngage' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
