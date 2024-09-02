import * as assert from 'assert';
import * as sinon from 'sinon';
import { cli } from '../../cli/cli';
import { Logger } from '../../cli/Logger';
import { telemetry } from '../../telemetry';
import { app } from '../../utils/app';
import { browserUtil } from '../../utils/browserUtil';
import { pid } from '../../utils/pid';
import { session } from '../../utils/session';
import { sinonUtil } from '../../utils/sinonUtil';
import commands from './commands';
import Command from '../../Command';
const command: Command = require('./docs');

describe(commands.DOCS, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let getSettingWithDefaultValueStub: sinon.SinonStub;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
    getSettingWithDefaultValueStub = sinon.stub(cli, 'getSettingWithDefaultValue').returns(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      loggerLogSpy,
      getSettingWithDefaultValueStub
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.DOCS);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should log a message and return if autoOpenLinksInBrowser is false', async () => {
    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(app.packageJson().homepage));
  });

  it('should open the CLI for Microsoft 365 docs webpage URL using "open" if autoOpenLinksInBrowser is true', async () => {
    getSettingWithDefaultValueStub.restore();
    getSettingWithDefaultValueStub = sinon.stub(cli, 'getSettingWithDefaultValue').returns(true);

    const openStub = sinon.stub(browserUtil, 'open').callsFake(async (url) => {
      if (url === 'https://pnp.github.io/cli-microsoft365/') {
        return;
      }
      throw 'Invalid url';
    });
    await command.action(logger, { options: {} });
    assert(openStub.calledWith(app.packageJson().homepage));
  });
});