<template>
  <div class="modal-overlay" @click.self="close">
    <div class="modal">
      <h2>Select Border Style</h2>
      <div class="options">
        <div class="border-grid">
          <button class="border-icon full" @click="apply('full')" title="Full border"></button>
          <button class="border-icon outer" @click="apply('outer')" title="Outer only"></button>
          <button class="border-icon inner" @click="apply('inner')" title="Inner only"></button>
          <button class="border-icon top" @click="apply('top')" title="Top only"></button>
          <button class="border-icon bottom" @click="apply('bottom')" title="Bottom only"></button>
          <button class="border-icon left" @click="apply('left')" title="Left only"></button>
          <button class="border-icon right" @click="apply('right')" title="Right only"></button>
          <button class="border-icon clear" @click="apply('clear')" title="Remove borders">✕</button>
        </div>
      </div>
      <button class="close-btn" @click="close">Close</button>
    </div>
  </div>
</template>

<script>
export default {
  name: 'BorderModal',
  emits: ['apply', 'close'],
  methods: {
    apply(type) {
      this.$emit('apply', type);  // ✅ emit instead of call
      this.$emit('close');
    },
    close() {
      this.$emit('close');
    }
  }
};
</script>

<style scoped>
.modal-overlay {
  position: fixed;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background: rgba(0, 0, 0, 0.5);
  display: flex;
  justify-content: center;
  align-items: center;
  z-index: 1000;
}
.modal {
  background: white;
  padding: 20px;
  border-radius: 8px;
  min-width: 300px;
  text-align: center;
}
.options {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  margin: 20px 0;
  justify-content: center;
}
.options button {
  padding: 6px 12px;
  background: #3498db;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
}
.options button:hover {
  background: #2980b9;
}
.close-btn {
  background: #ccc;
  padding: 6px 12px;
  border-radius: 4px;
  cursor: pointer;
}
.border-grid {
  display: grid;
  grid-template-columns: repeat(4, 40px);
  gap: 10px;
  justify-content: center;
  margin: 1rem 0;
}

.border-icon {
  width: 40px;
  height: 40px;
  background-color: #f4f4f4; /* 🔄 subtle gray for contrast */
  border: 2px solid #bbb;
  position: relative;
  cursor: pointer;
  box-shadow: inset 0 0 1px #999; /* 🧠 helps lines stand out */
}


.border-icon.full::after {
  content: "";
  position: absolute;
  inset: 2px;
border: 2px solid #1565c0; /* deeper blue for clarity */

}

.border-icon.outer::after {
  content: "";
  position: absolute;
  inset: 4px;
 border: 2px solid #1565c0; /* deeper blue for clarity */

}

.border-icon.inner::before,
.border-icon.inner::after {
  content: "";
  position: absolute;
 background-color: #1565c0;

}
.border-icon.inner::before {
  top: 50%;
  left: 5px;
  right: 5px;
  height: 2px;
  transform: translateY(-50%);
}
.border-icon.inner::after {
  left: 50%;
  top: 5px;
  bottom: 5px;
  width: 2px;
  transform: translateX(-50%);
}

.border-icon.top::after {
  content: "";
  position: absolute;
  top: 4px;
  left: 4px;
  right: 4px;
  height: 2px;
background-color: #1565c0;

}
.border-icon.bottom::after {
  content: "";
  position: absolute;
  bottom: 4px;
  left: 4px;
  right: 4px;
  height: 2px;
 background-color: #1565c0;

}
.border-icon.left::after {
  content: "";
  position: absolute;
  top: 4px;
  bottom: 4px;
  left: 4px;
  width: 2px;
background-color: #1565c0;

}
.border-icon.right::after {
  content: "";
  position: absolute;
  top: 4px;
  bottom: 4px;
  right: 4px;
  width: 2px;
background-color: #1565c0;

}

.border-icon.clear {
  font-size: 20px;
  background: #f8d7da;
  color: #a00;
  display: flex;
  align-items: center;
  justify-content: center;
  font-weight: bold;
}

.border-icon:hover {
  background-color: #e6f0ff;
  border-color: #3498db;
}

</style>
