import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../../Auth';
import { telemetry } from '../../telemetry';
import DelegatedGraphCommand from './DelegatedGraphCommand';
import { accessToken } from '../../utils/accessToken';
import { CommandError } from '../../Command';

class MockCommand extends DelegatedGraphCommand {
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

describe('DelegatedGraphCommand', () => {
  const cmd = new MockCommand();

  before(() => {
    sinon.stub(telemetry, 'trackEvent').returns();
    auth.connection.active = true;
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
  });

  beforeEach(() => {
    auth.connection.active = true;
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it(`doesn't throw error when not connected`, async () => {
    auth.connection.active = false;
    await (cmd as any).initAction({ options: {} }, {});
  });

  it('throws error when using application-only permissions', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    await assert.rejects(() => (cmd as any).initAction({ options: {} }, {}), new CommandError('This command does not support application-only permissions.'));
  });
});
