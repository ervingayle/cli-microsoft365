import * as assert from 'assert';
import * as sinon from 'sinon';
import { autocomplete } from '../../../../autocomplete';
import { Logger } from '../../../../cli/Logger';
import { telemetry } from '../../../../telemetry';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import commands from '../../commands';
import Command from '../../../../Command';
const command: Command = require('./completion-clink-update');

describe(commands.COMPLETION_CLINK_UPDATE, () => {
  let log: string[];
  let logger: Logger;
  let generateClinkCompletionStub: sinon.SinonStub;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    generateClinkCompletionStub = sinon.stub(autocomplete, 'getClinkCompletion').callsFake(() => '');
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
  });

  afterEach(() => {
    generateClinkCompletionStub.reset();
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.COMPLETION_CLINK_UPDATE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('builds command completion', async () => {
    await command.action(logger, { options: {} });
    assert(generateClinkCompletionStub.called);
  });
});
