import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import { telemetry } from '../../../../telemetry';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./accesstoken-get');

describe(commands.ACCESSTOKEN_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
      auth.ensureAccessToken
    ]);
    auth.connection.accessTokens = {};
    auth.connection.spoUrl = undefined;
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ACCESSTOKEN_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves access token for the specified resource', async () => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.connection.accessTokens['https://graph.microsoft.com'] = {
      expiresOn: d.toString(),
      accessToken: 'ABC'
    };

    await command.action(logger, { options: { resource: 'https://graph.microsoft.com' } });
    assert(loggerLogSpy.calledWith('ABC'));
  });

  it('retrieves access token for SharePoint when sharepoint specified as the resource and SPO URL previously retrieved', async () => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    auth.connection.accessTokens['https://contoso.sharepoint.com'] = {
      expiresOn: d.toString(),
      accessToken: 'ABC'
    };

    await command.action(logger, { options: { resource: 'sharepoint' } });
    assert(loggerLogSpy.calledWith('ABC'));
  });

  it('correctly handles error when retrieving access token', async () => {
    sinon.stub(auth, 'ensureAccessToken').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { resource: 'https://graph.microsoft.com' } } as any), new CommandError('An error has occurred'));
  });

  it('returns error when sharepoint specified as resource and SPO URL not available', async () => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.connection.accessTokens['https://contoso.sharepoint.com'] = {
      expiresOn: d.toString(),
      accessToken: 'ABC'
    };

    await assert.rejects(command.action(logger, { options: { resource: 'sharepoint' } } as any), new CommandError(`SharePoint URL undefined. Use the 'm365 spo set --url https://contoso.sharepoint.com' command to set the URL`));
  });

  it('retrieves access token for graph.microsoft.com when graph specified as the resource', async () => {
    const d: Date = new Date();
    d.setMinutes(d.getMinutes() + 1);
    auth.connection.accessTokens['https://graph.microsoft.com'] = {
      expiresOn: d.toString(),
      accessToken: 'ABC'
    };

    await command.action(logger, { options: { resource: 'graph' } });
    assert(loggerLogSpy.calledWith('ABC'));
  });
});
