import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
import Command, { CommandError } from '../../../../Command';
import { cli } from '../../../../cli/cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import request from '../../../../request';
import { telemetry } from '../../../../telemetry';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./retentioneventtype-get');

describe(commands.RETENTIONEVENTTYPE_GET, () => {
  const retentionEventTypeId = 'c37d695e-d581-4ae9-82a0-9364eba4291e';
  const retentionEventTypeGetResponse = {
    "displayName": "Test retention event type",
    "description": "Description for the retention event type",
    "createdDateTime": "2023-01-29T09:30:42Z",
    "lastModifiedDateTime": "2023-01-29T09:30:42Z",
    "id": "c37d695e-d581-4ae9-82a0-9364eba4291e",
    "createdBy": {
      "user": {
        "id": null,
        "displayName": "John Doe"
      }
    },
    "lastModifiedBy": {
      "user": {
        "id": null,
        "displayName": "John Doe"
      }
    }
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
    auth.connection.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.RETENTIONEVENTTYPE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if a correct id is entered', async () => {
    const actual = await command.validate({ options: { id: retentionEventTypeId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves retention event type by specified id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggerTypes/retentionEventTypes/${retentionEventTypeId}`) {
        return retentionEventTypeGetResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: retentionEventTypeId, verbose: true } });
    assert(loggerLogSpy.calledWith(retentionEventTypeGetResponse));
  });

  it('handles error when retention event type by specified id is not found', async () => {
    const errorMessage = `Error: The operation couldn't be performed because object '${retentionEventTypeId}' couldn't be found on 'FfoConfigurationSession'.`;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggerTypes/retentionEventTypes/${retentionEventTypeId}`) {
        throw errorMessage;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: retentionEventTypeId
      }
    }), new CommandError(errorMessage));
  });
});