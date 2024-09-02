import * as assert from 'assert';
import * as sinon from "sinon";
import auth from '../../../../Auth';
import { CommandInfo } from "../../../../cli/CommandInfo";
import { Logger } from "../../../../cli/Logger";
import commands from "../../commands";
import { telemetry } from '../../../../telemetry';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { cli } from '../../../../cli/cli';
const command: Command = require('./administrativeunit-get');
import request from '../../../../request';
import { sinonUtil } from '../../../../utils/sinonUtil';
import Command, { CommandError } from '../../../../Command';
import { entraAdministrativeUnit } from '../../../../utils/entraAdministrativeUnit';
import aadCommands from '../../aadCommands';

describe(commands.ADMINISTRATIVEUNIT_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const administrativeUnitsReponse = {
    value: [
      {
        id: 'fc33aa61-cf0e-46b6-9506-f633347202ab',
        displayName: 'European Division',
        visibility: 'HiddenMembership'
      },
      {
        id: 'a25b4c5e-e8b7-4f02-a23d-0965b6415098',
        displayName: 'Asian Division',
        visibility: null
      }
    ]
  };
  const validId = 'fc33aa61-cf0e-46b6-9506-f633347202ab';
  const validDisplayName = 'European Division';

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      entraAdministrativeUnit.getAdministrativeUnitByDisplayName,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ADMINISTRATIVEUNIT_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.deepStrictEqual(alias, [aadCommands.ADMINISTRATIVEUNIT_GET]);
  });

  it('retrieves information about the specified administrative unit by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/directory/administrativeUnits/${validId}`) {
        return administrativeUnitsReponse.value[0];
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: validId } });
    assert(loggerLogSpy.calledOnceWithExactly(administrativeUnitsReponse.value[0]));
  });

  it('retrieves information about the specified administrative unit by displayName', async () => {
    sinon.stub(entraAdministrativeUnit, 'getAdministrativeUnitByDisplayName').resolves(administrativeUnitsReponse.value[0]);

    await command.action(logger, { options: { displayName: validDisplayName } });
    assert(loggerLogSpy.calledOnceWithExactly(administrativeUnitsReponse.value[0]));
  });

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects(new Error(errorMessage));

    await assert.rejects(command.action(logger, { options: { id: validId } }), new CommandError(errorMessage));
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: validId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (displayName)', async () => {
    const actual = await command.validate({ options: { displayName: validDisplayName } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});