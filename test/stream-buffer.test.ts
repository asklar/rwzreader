import { describe, it, expect } from 'vitest';
import { StreamBuffer } from '../src/stream-buffer.js';
import { u8, u16, u32 } from './helpers.js';

describe('StreamBuffer', () => {
  it('readUInt8 reads a byte and advances offset', () => {
    const sb = new StreamBuffer(Buffer.from([0x42, 0xff]));
    expect(sb.readUInt8()).toBe(0x42);
    expect(sb.offset).toBe(1);
    expect(sb.readUInt8()).toBe(0xff);
    expect(sb.offset).toBe(2);
  });

  it('readUInt16 reads little-endian uint16', () => {
    const buf = Buffer.alloc(4);
    buf.writeUInt16LE(0x1234, 0);
    buf.writeUInt16LE(0xabcd, 2);
    const sb = new StreamBuffer(buf);
    expect(sb.readUInt16()).toBe(0x1234);
    expect(sb.readUInt16()).toBe(0xabcd);
    expect(sb.offset).toBe(4);
  });

  it('readUInt32 reads little-endian uint32', () => {
    const sb = new StreamBuffer(u32(0xdeadbeef));
    expect(sb.readUInt32()).toBe(0xdeadbeef);
    expect(sb.offset).toBe(4);
  });

  it('readUInt64 reads little-endian uint64 as BigInt', () => {
    const buf = Buffer.alloc(8);
    buf.writeBigUInt64LE(0x123456789abcdef0n);
    const sb = new StreamBuffer(buf);
    expect(sb.readUInt64()).toBe(0x123456789abcdef0n);
    expect(sb.offset).toBe(8);
  });

  it('readDouble reads little-endian double', () => {
    const buf = Buffer.alloc(8);
    buf.writeDoubleLE(3.14);
    const sb = new StreamBuffer(buf);
    expect(sb.readDouble()).toBeCloseTo(3.14);
    expect(sb.offset).toBe(8);
  });

  it('readAsciiString reads N ASCII bytes', () => {
    const sb = new StreamBuffer(Buffer.from('Hello\0World'));
    expect(sb.readAsciiString(5)).toBe('Hello');
    expect(sb.offset).toBe(5);
  });

  it('readAsciiUntilNullTerminator stops at 0x00', () => {
    const sb = new StreamBuffer(Buffer.from('test\0extra'));
    expect(sb.readAsciiUntilNullTerminator()).toBe('test');
  });

  it('readStringObject reads short string (length < 0xFF)', () => {
    // length=3, then 3 UTF-16LE chars "abc"
    const parts = [u8(3), u16(0x61), u16(0x62), u16(0x63)];
    const sb = new StreamBuffer(Buffer.concat(parts));
    expect(sb.readStringObject()).toBe('abc');
  });

  it('readBytes returns correct slice and advances offset', () => {
    const sb = new StreamBuffer(Buffer.from([1, 2, 3, 4, 5]));
    const bytes = sb.readBytes(3);
    expect(bytes).toEqual(Buffer.from([1, 2, 3]));
    expect(sb.offset).toBe(3);
  });

  it('readBytes throws on buffer underrun', () => {
    const sb = new StreamBuffer(Buffer.alloc(4));
    expect(() => sb.readBytes(5)).toThrow('buffer underrun');
  });

  it('readBytes throws on negative length', () => {
    const sb = new StreamBuffer(Buffer.alloc(4));
    expect(() => sb.readBytes(-1)).toThrow('buffer underrun');
  });

  it('remaining decreases as data is read', () => {
    const sb = new StreamBuffer(Buffer.alloc(10));
    expect(sb.remaining).toBe(10);
    sb.readUInt32();
    expect(sb.remaining).toBe(6);
    sb.readUInt16();
    expect(sb.remaining).toBe(4);
  });

  it('chained reads produce correct final offset', () => {
    const buf = Buffer.alloc(15);
    const sb = new StreamBuffer(buf);
    sb.readUInt32(); // 4
    sb.readUInt16(); // 6
    sb.readUInt8();  // 7
    sb.readDouble(); // 15
    expect(sb.offset).toBe(15);
    expect(sb.remaining).toBe(0);
  });
});
