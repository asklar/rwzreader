export class StreamBuffer {
  public offset = 0;

  public constructor(private readonly buffer: Buffer) {}

  public readUInt64() {
    const v = this.buffer.readBigUInt64LE(this.offset);
    this.offset += 8;
    return v;
  }

  public readUInt32() {
    const v = this.buffer.readUInt32LE(this.offset);
    this.offset += 4;
    return v;
  }

  public readUInt16() {
    const v = this.buffer.readUInt16LE(this.offset);
    this.offset += 2;
    return v;
  }

  public readUInt8() {
    const v = this.buffer.readUInt8(this.offset);
    this.offset++;
    return v;
  }

  public readDouble() {
    const v = this.buffer.readDoubleLE(this.offset);
    this.offset += 8;
    return v;
  }

  public readString(len: number) {
    const b = Buffer.alloc(len);
    for (let i = 0; i < len; i++) {
      b[i] = this.readUInt16();
    }
    return b.toString();
  }

  public readStringObject() {
    let length = this.readUInt8();
    if (length === 0xff) {
      length = this.readUInt16();
      this.offset += 2;
    }
    return this.readString(length);
  }

  public readAsciiString(len: number) {
    const b = Buffer.alloc(len);
    try {
      for (let i = 0; i < len; i++) {
        b[i] = this.readUInt8();
      }
      return b.toString().replace(/\0+$/, '');
    } catch (e) {
      return '';
    }
  }

  public readStringUntilNullTerminator() {
    const v = Buffer.alloc(260);
    try {
      let i = 0;
      for (let c = this.readUInt16(); c !== 0; c = this.readUInt16()) {
        v.writeUInt16LE(c, i++);
      }
      return v.toString('utf8', 0, i);
    } catch (e) {
      return '';
    }
  }

  public readAsciiUntilNullTerminator() {
    const v = Buffer.alloc(260);
    let i = 0;
    for (let c = this.readUInt8(); c !== 0; c = this.readUInt8()) {
      v.writeUInt16LE(c, i++);
    }
    return v.toString('utf8', 0, i);
  }

  public readBytes(len: number): Buffer {
    if (len < 0 || this.offset + len > this.buffer.length) {
      throw new Error(`buffer underrun: requested ${len} bytes at offset ${this.offset}, buffer length ${this.buffer.length}`);
    }
    const b = this.buffer.subarray(this.offset, this.offset + len);
    this.offset += len;
    return Buffer.from(b as any);
  }

  public get remaining() {
    return this.buffer.length - this.offset;
  }
}
