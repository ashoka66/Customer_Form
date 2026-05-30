const mongoose = require("mongoose");
const bcrypt   = require("bcryptjs");

/*
  User model

  Purpose and runtime notes:
  - Represents application users (both `superadmin` and `admin`).
  - Used by `server.js` for authentication, session storage (only selected fields), and admin management.
  - Fields: `name`, `email`, `password`, `role`, `isActive`, `createdBy`, `createdByName`, `lastLogin`.
  - Pre-save hook hashes the password whenever it is modified before saving to the database.
  - `comparePassword` is an instance method used during login to verify credentials.
*/

const UserSchema = new mongoose.Schema({
  name:          { type: String, required: true },
  email:         { type: String, required: true, unique: true, lowercase: true, trim: true },
  password:      { type: String, required: true },
  role:          { type: String, enum: ["superadmin", "admin"], default: "admin" },
  isActive:      { type: Boolean, default: true },
  createdBy:     { type: String, default: "" },
  createdByName: { type: String, default: "" },
  lastLogin:     { type: Date, default: null },
}, { timestamps: true });

// ✅ No next() — returns a promise directly, works in all Node versions
UserSchema.pre("save", async function() {
  if (!this.isModified("password")) return;
  this.password = await bcrypt.hash(this.password, 10);
});

UserSchema.methods.comparePassword = async function(enteredPassword) {
  return await bcrypt.compare(enteredPassword, this.password);
};

module.exports = mongoose.model("User", UserSchema);