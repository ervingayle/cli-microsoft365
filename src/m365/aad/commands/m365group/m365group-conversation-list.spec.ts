import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./m365group-conversation-list');
import { aadGroup } from '../../../../utils/aadGroup';

describe(commands.M365GROUP_CONVERSATION_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const jsonOutput = {
    "value": [
      {
        "id": "AAQkAGFhZDhkNGI1LTliZmEtNGEzMi04NTkzLWZjMWExZDkyMWEyZgAQAH4o7SknOTNKqAqMhqJHtUM=",
        "topic": "The new All Company group is ready",
        "hasAttachments": false,
        "lastDeliveredDateTime": "2021-08-02T10:34:00Z",
        "uniqueSenders": [
          "All Company"
        ],
        "preview": "Welcome to the All Company group.Use the group to share ideas, files, and important dates.Start a conversationRead group conversations or start your own.Share filesView, edit, and share all group files, including email attachments.Connect your"
      },
      {
        "id": "AAQkADQzYWUxZTA5LWQwYmItNDcxMy04ZTU5LTg3YmU5NDU3MDlmZgAQAOAzGByAP-dGkFXbXlukzUk=",
        "topic": "Weekly meeting on friday",
        "hasAttachments": false,
        "lastDeliveredDateTime": "2022-02-02T10:34:00Z",
        "uniqueSenders": [
          "John Doe"
        ],
        "preview": "Can we have a weekly meeting on Friday?"
      }
    ]
  };
  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(aadGroup, 'isUnifiedGroup').resolves(true);
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
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
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.M365GROUP_CONVERSATION_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['topic', 'lastDeliveredDateTime', 'id']);
  });
  it('fails validation if the groupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { groupId: 'not-c49b-4fd4-8223-28f0ac3a6402' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
  it('passes validation if the groupId is a valid GUID', async () => {
    const actual = await command.validate({ options: { groupId: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('Retrieve conversations for the specified group by groupId in the tenant (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/conversations`) {
        return jsonOutput;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        verbose: true, groupId: "00000000-0000-0000-0000-000000000000"
      }
    });
    assert(loggerLogSpy.calledWith(
      jsonOutput.value
    ));
  });
  it('correctly handles error when listing conversations', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { groupId: "00000000-0000-0000-0000-000000000000" } } as any),
      new CommandError('An error has occurred'));
  });

  it('shows error when the group is not a unified group', async () => {
    const groupId = '3f04e370-cbc6-4091-80fe-1d038be2ad06';

    sinonUtil.restore(aadGroup.isUnifiedGroup);
    sinon.stub(aadGroup, 'isUnifiedGroup').resolves(false);

    await assert.rejects(command.action(logger, { options: { groupId: groupId } } as any),
      new CommandError(`Specified group with id '${groupId}' is not a Microsoft 365 group.`));
  });

});
