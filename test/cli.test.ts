import { describe, it, expect } from 'vitest';
import { execFileSync } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { buildMinimalRwz } from './helpers.js';

const CLI = path.resolve('bin/index.js');

function runCli(args: string[]): { stdout: string; stderr: string; exitCode: number } {
  try {
    const stdout = execFileSync('node', [CLI, ...args], {
      encoding: 'utf-8',
      timeout: 10000,
    });
    return { stdout, stderr: '', exitCode: 0 };
  } catch (e: any) {
    return {
      stdout: e.stdout ?? '',
      stderr: e.stderr ?? '',
      exitCode: e.status ?? 1,
    };
  }
}

describe('CLI', () => {
  it('prints usage when no arguments provided', () => {
    const { stderr, exitCode } = runCli([]);
    expect(exitCode).not.toBe(0);
    expect(stderr).toContain('Usage');
  });

  it('errors when input file does not exist', () => {
    const { stderr, exitCode } = runCli(['nonexistent.rwz']);
    expect(exitCode).not.toBe(0);
    expect(stderr).toContain('File not found');
  });

  it('converts a valid rwz file to JSON', () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'rwztest-'));
    const inputFile = path.join(tmpDir, 'test.rwz');
    const outputFile = path.join(tmpDir, 'output.json');

    try {
      fs.writeFileSync(inputFile, buildMinimalRwz(2));
      const { stdout, exitCode } = runCli([inputFile, outputFile]);
      expect(exitCode).toBe(0);
      expect(stdout).toContain('Converted');

      const json = JSON.parse(fs.readFileSync(outputFile, 'utf-8'));
      expect(json.rules).toHaveLength(2);
      expect(json.header.version).toBe('outlook2019');
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  it('uses default output path when not specified', () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'rwztest-'));
    const inputFile = path.join(tmpDir, 'test.rwz');
    const expectedOutput = path.join(tmpDir, 'outlook-rules.json');

    try {
      fs.writeFileSync(inputFile, buildMinimalRwz(1));
      // Run from tmpDir so default output lands there
      const stdout = execFileSync('node', [CLI, inputFile], {
        encoding: 'utf-8',
        cwd: tmpDir,
        timeout: 10000,
      });
      expect(stdout).toContain('Converted');
      // The default output is 'outlook-rules.json' in cwd
    } finally {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });
});
