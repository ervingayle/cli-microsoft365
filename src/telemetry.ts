import * as child_process from 'child_process';
import * as path from 'path';
import { cli } from './cli/cli';
import { settingsNames } from './settingsNames';
import { pid } from './utils/pid';
import { session } from './utils/session';

function trackTelemetry(object: any): void {
  try {
    const child = child_process.spawn('node', [path.join(__dirname, 'telemetryRunner')], {
      stdio: ['pipe', 'ignore', 'ignore'],
      detached: true
    });
    child.unref();

    object.shell = pid.getProcessName(process.ppid) || '';
    object.session = session.getId(process.ppid);

    child.stdin.write(JSON.stringify(object));
    child.stdin.end();
  }
  catch { }
}

export const telemetry = {
  trackEvent: (commandName: string, properties: any): void => {
    if (cli.getSettingWithDefaultValue<boolean>(settingsNames.disableTelemetry, false)) {
      return;
    }

    trackTelemetry({
      commandName,
      properties
    });
  }
};