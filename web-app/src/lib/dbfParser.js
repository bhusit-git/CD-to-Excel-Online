import iconv from 'iconv-lite';
import { Buffer } from 'buffer';

/**
 * Parses a DBF file from an ArrayBuffer
 * @param {ArrayBuffer} buffer - The file content as ArrayBuffer
 * @returns {Array<Object>} - An array of records (dictionaries)
 */
export function parseDbf(buffer) {
  const dataView = new DataView(buffer);
  const uint8Array = new Uint8Array(buffer);
  
  // Read Header
  // Number of records
  const numRecords = dataView.getUint32(4, true); // true for little-endian
  // Length of header structure
  const headerLen = dataView.getUint16(8, true);
  // Length of each record
  const recordLen = dataView.getUint16(10, true);
  
  // Read Field Descriptors
  const fields = [];
  let offset = 32;
  
  while (offset < headerLen) {
    // Check for end of header marker
    if (uint8Array[offset] === 0x0D) {
      break;
    }
    
    // Read field name (up to 11 bytes, null-terminated)
    let nameEnd = 0;
    while (nameEnd < 11 && uint8Array[offset + nameEnd] !== 0x00) {
      nameEnd++;
    }
    const nameBytes = new Uint8Array(buffer, offset, nameEnd);
    const name = new TextDecoder('ascii').decode(nameBytes).trim();
    
    // Field type
    const type = String.fromCharCode(uint8Array[offset + 11]);
    
    // Field length
    const len = uint8Array[offset + 16];
    
    fields.push({ name, type, len });
    offset += 32;
  }
  
  // Calculate field offsets within a record
  const fieldOffsets = [];
  let currentOffset = 1; // 1 byte for deletion flag
  for (const field of fields) {
    fieldOffsets.push(currentOffset);
    currentOffset += field.len;
  }
  
  // Read Records
  const records = [];
  let recordStart = headerLen;
  
  for (let i = 0; i < numRecords; i++) {
    if (recordStart + recordLen > buffer.byteLength) {
      break; // Reached end of buffer unexpectedly
    }
    
    const deletionFlag = String.fromCharCode(uint8Array[recordStart]);
    if (deletionFlag !== '*') {
      // Record is not deleted
      const record = {};
      
      for (let j = 0; j < fields.length; j++) {
        const field = fields[j];
        const fieldOffset = recordStart + fieldOffsets[j];
        const fieldData = new Uint8Array(buffer, fieldOffset, field.len);
        
        // Remove trailing null bytes and spaces before decoding to save time
        let endIdx = field.len;
        while (endIdx > 0 && (fieldData[endIdx - 1] === 0x00 || fieldData[endIdx - 1] === 0x20)) {
          endIdx--;
        }
        
        const cleanData = fieldData.slice(0, endIdx);
        
        let value;
        if (field.type === 'D') {
          value = new TextDecoder('ascii').decode(cleanData).trim();
        } else {
          // Decode CP874 (Thai)
          value = iconv.decode(Buffer.from(cleanData), 'cp874').trim();
        }
        
        record[field.name] = value;
      }
      records.push(record);
    }
    
    recordStart += recordLen;
  }
  
  return records;
}
