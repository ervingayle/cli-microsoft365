import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../telemetry';
import auth from '../../Auth';
import { Logger } from '../../cli/Logger';
import { CommandError } from '../../Command';
import { pid } from '../../utils/pid';
import { session } from '../../utils/session';
import { sinonUtil } from '../../utils/sinonUtil';
import PowerBICommand from './PowerBICommand';
import { accessToken } from '../../utils/accessToken';

class MockCommand extends PowerBICommand {
  public get name(): string {
    return 'mock';
  }

  public get description(): string {
    return 'Mock command';
  }

  public async commandAction(): Promise<void> {
  }

  public commandHelp(): void {
  }
}

describe('PowerBICommand', () => {
  before(() => {
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
  });

  afterEach(() => {
    sinonUtil.restore(auth.restoreAuth);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('correctly reports an error while restoring auth info', async () => {
    sinon.stub(auth, 'restoreAuth').callsFake(async () => { throw 'An error has occurred'; });
    const command = new MockCommand();
    const logger: Logger = {
      log: async () => { },
      logRaw: async () => { },
      logToStderr: async () => { }
    };
    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });

  it('doesn\'t execute command when error occurred while restoring auth info', async () => {
    sinon.stub(auth, 'restoreAuth').callsFake(async () => { throw 'An error has occurred'; });
    const command = new MockCommand();
    const logger: Logger = {
      log: async () => { },
      logRaw: async () => { },
      logToStderr: async () => { }
    };
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    await assert.rejects(command.action(logger, { options: {} }));
    assert(commandCommandActionSpy.notCalled);
  });

  it('doesn\'t execute command when not logged in', async () => {
    sinon.stub(auth, 'restoreAuth').resolves();
    const command = new MockCommand();
    const logger: Logger = {
      log: async () => { },
      logRaw: async () => { },
      logToStderr: async () => { }
    };
    auth.connection.active = false;
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    await assert.rejects(command.action(logger, { options: {} }));
    assert(commandCommandActionSpy.notCalled);
  });

  it('executes command when logged in', async () => {
    sinon.stub(auth, 'restoreAuth').resolves();
    const command = new MockCommand();
    const logger: Logger = {
      log: async () => { },
      logRaw: async () => { },
      logToStderr: async () => { }
    };
    auth.connection.active = true;
    const commandCommandActionSpy = sinon.spy(command, 'commandAction');
    await command.action(logger, { options: {} });
    assert(commandCommandActionSpy.called);
  });

  it('returns correct api resource', () => {
    const command = new MockCommand();
    assert.strictEqual((command as any).resource, 'https://api.powerbi.com');
  });

  it('throws error when using application-only permissions', async () => {
    const cmd = new MockCommand();
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    auth.connection.active = true;
    await assert.rejects(() => (cmd as any).initAction({ options: {} }, {}), new CommandError('This command does not support application-only permissions.'));
  });
});
