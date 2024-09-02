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
import type * as Chalk from 'chalk';
const command: Command = require('./engage-message-remove');
import { settingsNames } from '../../../../settingsNames';
import yammerCommands from './yammerCommands';
import { accessToken } from '../../../../utils/accessToken';

describe(commands.ENGAGE_MESSAGE_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertDelegatedAccessToken').returns();
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      cli.promptForConfirmation,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ENGAGE_MESSAGE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.deepStrictEqual(alias, [yammerCommands.MESSAGE_REMOVE]);
  });

  it('correctly logs deprecation warning for yammer command', async () => {
    const chalk: typeof Chalk = require('chalk');
    const loggerErrSpy = sinon.spy(logger, 'logToStderr');
    const commandNameStub = sinon.stub(cli, 'currentCommandName').value(yammerCommands.MESSAGE_REMOVE);
    sinon.stub(request, 'delete').resolves();

    await command.action(logger, { options: { id: 10123190123123, force: true } });
    assert.deepStrictEqual(loggerErrSpy.firstCall.firstArg, chalk.yellow(`Command '${yammerCommands.MESSAGE_REMOVE}' is deprecated. Please use '${commands.ENGAGE_MESSAGE_REMOVE}' instead.`));

    sinonUtil.restore([loggerErrSpy, commandNameStub]);
  });

  it('id must be a number', async () => {
    const actual = await command.validate({ options: { id: 'nonumber' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('id is required', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('calls the messaging endpoint with the right parameters and confirmation', async () => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/10123190123123.json') {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: 10123190123123, force: true } });
    assert.strictEqual(requestDeleteStub.lastCall.args[0].url, 'https://www.yammer.com/api/v1/messages/10123190123123.json');
  });

  it('calls the messaging endpoint with the right parameters without confirmation', async () => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/10123190123123.json') {
        return;
      }
      throw 'Invalid request';
    });
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { debug: true, id: 10123190123123, force: false } });
    assert.strictEqual(requestDeleteStub.lastCall.args[0].url, 'https://www.yammer.com/api/v1/messages/10123190123123.json');
  });

  it('does not call the messaging endpoint without confirmation', async () => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/10123190123123.json') {
        return;
      }
      throw 'Invalid request';
    });

    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { debug: true, id: 10123190123123, force: false } });
    assert(requestDeleteStub.notCalled);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'delete').callsFake(async () => {
      throw {
        "error": {
          "base": "An error has occurred."
        }
      };
    });

    await assert.rejects(command.action(logger, { options: { id: 10123190123123, force: true } } as any), new CommandError('An error has occurred.'));
  });

  it('passes validation with parameters', async () => {
    const actual = await command.validate({ options: { id: 10123123 } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
